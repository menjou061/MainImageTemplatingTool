param(
    [string]$ExcelPath,
    [string]$PsdPath,
    [string]$OutputRoot,
    [string]$SheetName,
    [string]$ProductName,
    [int]$Limit,
    [switch]$NoUi
)

$ErrorActionPreference = 'Stop'

# Sheet names are exchanged with the bundled Python process. Force the pipe to
# UTF-8 so Chinese worksheet names survive the Python -> PowerShell -> Python round trip.
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[Console]::InputEncoding = $utf8NoBom
[Console]::OutputEncoding = $utf8NoBom
$OutputEncoding = $utf8NoBom
$env:PYTHONIOENCODING = 'utf-8'

$script:LogLines = New-Object System.Collections.Generic.List[string]
$script:OpenedDocument = $null
$script:Stage = '入口初始化'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:InternalDir = $scriptDir
$baseDir = Split-Path -Parent $scriptDir
$userDataRoot = if (-not [string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) {
    Join-Path $env:LOCALAPPDATA '电商主图套版工具'
} else {
    Join-Path $baseDir '_user_data'
}
$diagnosticDir = Join-Path $userDataRoot '任务记录'
$entryReportTxt = Join-Path $diagnosticDir 'failure.txt'
$entryReportCsv = Join-Path $diagnosticDir 'failure.csv'
$entryStatusJson = Join-Path $diagnosticDir 'status.json'
$entryStatusTxt = Join-Path $diagnosticDir 'status.txt'
$runtimeProbeTxt = Join-Path $diagnosticDir 'runtime_probe.txt'
$settingsPath = Join-Path $userDataRoot 'settings.json'
$taskHistoryPath = Join-Path $diagnosticDir '任务历史.csv'
$ps1Marker = Join-Path $diagnosticDir 'ps1_started.marker'
$cleanScript = Join-Path $scriptDir 'clean_data.py'
$sheetScript = Join-Path $scriptDir 'l0_list_sheets.py'
$batchScript = Join-Path $scriptDir 'batch_template.jsx'
$templatePrepareScript = Join-Path $scriptDir 'template_prepare.jsx'
$runtimeRoot = Join-Path $scriptDir 'runtime'
$privatePythonDir = Join-Path $runtimeRoot 'python'
$privatePythonExe = Join-Path $privatePythonDir 'python.exe'
$runtimeTempDir = Join-Path $runtimeRoot '_download_tmp'
$pythonEmbedUrl = 'https://www.python.org/ftp/python/3.11.9/python-3.11.9-embed-amd64.zip'
$getPipUrl = 'https://bootstrap.pypa.io/get-pip.py'
$requiredPythonVersion = '3.11.9'
$requiredOpenpyxlVersion = '3.1.5'
$requiredEtXmlfileVersion = '2.0.0'
$taskOutputDir = $null
$taskRecordDir = $null
$reportPath = $null
$script:Settings = $null
$script:ProgressForm = $null
$script:ProgressStageLabel = $null
$script:ProgressDetailLabel = $null
$script:LoadedTaskSettingsExcelPath = ''
$generationStartedAt = $null
$script:CurrentExcelPath = ''
$script:CurrentPsdPath = ''
$script:CurrentSheetName = ''
$script:CurrentProductName = ''

function Write-Utf8Bom {
    param(
        [string]$Path,
        [string]$Content
    )
    $encoding = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($Path, $Content, $encoding)
}

function Add-Log {
    param([string]$Message)
    $timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $script:LogLines.Add("[$timestamp] $Message") | Out-Null
    Write-Host $Message
}

function Write-EntryStatus {
    param(
        [string]$Status,
        [string]$Message
    )
    $payload = [pscustomobject]@{
        status = $Status
        time = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        stage = $script:Stage
        message = $Message
        exitCode = if ($Status -eq 'success' -or $Status -eq 'needs_review') { 0 } elseif ($Status -eq 'running') { $null } else { 1 }
    }
    $json = $payload | ConvertTo-Json -Depth 4
    Write-Utf8Bom -Path $entryStatusJson -Content ($json + [Environment]::NewLine)
    Write-Utf8Bom -Path $entryStatusTxt -Content ("状态：$Status`r`n时间：$($payload.time)`r`n当前阶段：$($script:Stage)`r`n退出状态：$($payload.exitCode)`r`n说明：$Message`r`n")
}

function Write-EntryFailureReport {
    param(
        [string]$ErrorSummary,
        [string]$Suggestion
    )
    $timeText = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $logText = $script:LogLines -join [Environment]::NewLine
    $taskOutputText = '未建立'
    if ($taskOutputDir) {
        $taskOutputText = $taskOutputDir
    }
    $txt = @(
        'L0 电商主图批量套版入口失败报告',
        '',
        '状态：未通过',
        "时间：$timeText",
        "当前阶段：$script:Stage",
        '退出状态：失败',
        "包目录：$baseDir",
        "内部脚本目录：$scriptDir",
        "固定诊断目录：$diagnosticDir",
        "任务输出目录：$taskOutputText",
        "错误摘要：$ErrorSummary",
        "建议动作：$Suggestion",
        '',
        '运行日志：',
        $logText
    ) -join [Environment]::NewLine
    Write-Utf8Bom -Path $entryReportTxt -Content ($txt + [Environment]::NewLine)

    $csv = @(
        '"字段","值"',
        '"状态","未通过"',
        ('"时间","{0}"' -f ($timeText -replace '"', '""')),
        ('"当前阶段","{0}"' -f ($script:Stage -replace '"', '""')),
        '"退出状态","失败"',
        ('"包目录","{0}"' -f ($baseDir -replace '"', '""')),
        ('"内部脚本目录","{0}"' -f ($scriptDir -replace '"', '""')),
        ('"固定诊断目录","{0}"' -f ($diagnosticDir -replace '"', '""')),
        ('"任务输出目录","{0}"' -f ($taskOutputText -replace '"', '""')),
        ('"错误摘要","{0}"' -f ($ErrorSummary -replace '"', '""')),
        ('"建议动作","{0}"' -f ($Suggestion -replace '"', '""'))
    ) -join [Environment]::NewLine
    Write-Utf8Bom -Path $entryReportCsv -Content ($csv + [Environment]::NewLine)
    Write-EntryStatus -Status 'failed' -Message $ErrorSummary
}

function Get-DefaultOutputRoot {
    $desktop = [Environment]::GetFolderPath([Environment+SpecialFolder]::DesktopDirectory)
    if ([string]::IsNullOrWhiteSpace($desktop)) {
        $desktop = [Environment]::GetFolderPath([Environment+SpecialFolder]::MyPictures)
    }
    if ([string]::IsNullOrWhiteSpace($desktop)) {
        $desktop = $baseDir
    }
    return (Join-Path $desktop '电商主图套版成品')
}

function ConvertTo-SafePathPart {
    param([string]$Value)
    $text = if ([string]::IsNullOrWhiteSpace($Value)) { '全部商品' } else { $Value.Trim() }
    foreach ($character in [System.IO.Path]::GetInvalidFileNameChars()) {
        $text = $text.Replace([string]$character, '_')
    }
    $text = $text -replace '\s+', '_'
    if ($text.Length -gt 40) {
        $text = $text.Substring(0, 40)
    }
    $text = $text.Trim(' ', '.', '_')
    if ([string]::IsNullOrWhiteSpace($text)) { return '任务' }
    return $text
}

function New-TaskOutputDirectory {
    param(
        [string]$OutputRoot,
        [string]$SheetName
    )
    $baseName = '套版成品_{0}_{1}' -f (Get-Date).ToString('yyyy-MM-dd_HHmmss'), (ConvertTo-SafePathPart $SheetName)
    $candidate = Join-Path $OutputRoot $baseName
    $suffix = 1
    while (Test-Path -LiteralPath $candidate) {
        $candidate = Join-Path $OutputRoot ("{0}_{1}" -f $baseName, $suffix)
        $suffix++
    }
    New-Item -Path $candidate -ItemType Directory -Force | Out-Null
    return $candidate
}

function Write-TaskHistory {
    param(
        [string]$Status,
        [string]$Message
    )
    try {
        New-Item -Path $diagnosticDir -ItemType Directory -Force | Out-Null
        $row = [pscustomobject][ordered]@{
            '时间' = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            '状态' = $Status
            '当前阶段' = $script:Stage
            '工作表' = $script:CurrentSheetName
            '商品范围' = if ([string]::IsNullOrWhiteSpace($script:CurrentProductName)) { '全部商品' } else { $script:CurrentProductName }
            '商品表格' = $script:CurrentExcelPath
            'PSD模板' = $script:CurrentPsdPath
            '任务文件夹' = if ($taskOutputDir) { $taskOutputDir } else { '' }
            '说明' = $Message
        }
        if (Test-Path -LiteralPath $taskHistoryPath -PathType Leaf) {
            $row | Export-Csv -LiteralPath $taskHistoryPath -NoTypeInformation -Encoding UTF8 -Append
        } else {
            $row | Export-Csv -LiteralPath $taskHistoryPath -NoTypeInformation -Encoding UTF8
        }
    } catch {
        Add-Log "写入任务历史失败：$($_.Exception.Message)"
    }
}

function Clear-StaleDiagnostics {
    foreach ($path in @($entryReportTxt, $entryReportCsv, $runtimeProbeTxt)) {
        try {
            if (Test-Path -LiteralPath $path -PathType Leaf) {
                Remove-Item -LiteralPath $path -Force -ErrorAction Stop
                Add-Log "已清理旧诊断文件：$path"
            }
        } catch {
            Add-Log "清理旧诊断文件失败，将继续运行：$path；$($_.Exception.Message)"
        }
    }
}

function Get-PhotoshopResultSummary {
    param(
        [string]$ResultReport,
        [string]$JpgOutputDir
    )
    if (-not (Test-Path -LiteralPath $ResultReport -PathType Leaf)) {
        throw "未找到 Photoshop 结果报告：$ResultReport"
    }

    $rows = @(Import-Csv -LiteralPath $ResultReport -Encoding UTF8)
    $criticalCodes = @('E_TEMPLATE_INVALID', 'E_MISSING_IMAGE', 'E_PRICE_INVALID', 'E_EMPTY_FIELD')
    $criticalRows = New-Object System.Collections.Generic.List[object]
    $failedRows = New-Object System.Collections.Generic.List[object]
    $reviewRows = New-Object System.Collections.Generic.List[object]
    $successRows = New-Object System.Collections.Generic.List[object]
    foreach ($row in $rows) {
        $status = [string]$row.状态
        $codes = [string]$row.错误码
        $hasOutput = -not [string]::IsNullOrWhiteSpace([string]$row.输出文件)
        $hasCriticalCode = $false
        foreach ($code in $criticalCodes) {
            if ($codes -match "(^|;)$([regex]::Escape($code))(;|$)") {
                $hasCriticalCode = $true
                break
            }
        }
        # A row with an actual JPG is not an export failure.  Field/image
        # warnings remain visible as "需复核" so the designer can decide
        # whether the generated result is usable.
        if (-not $hasOutput -or $status -eq '处理失败' -or $status -eq '模板错误' -or $status -eq '数据需核对') {
            $failedRows.Add($row) | Out-Null
        } elseif ($status -eq '成功') {
            $successRows.Add($row) | Out-Null
        } else {
            $reviewRows.Add($row) | Out-Null
        }
        if ($hasCriticalCode -and -not $hasOutput) {
            $criticalRows.Add($row) | Out-Null
        }
    }

    $exportedCount = @(Get-ChildItem -LiteralPath $JpgOutputDir -Filter '*.jpg' -File -ErrorAction SilentlyContinue).Count
    $outcome = 'success'
    if ($rows.Count -eq 0 -or ($exportedCount -eq 0 -and $failedRows.Count -eq 0 -and $reviewRows.Count -eq 0)) {
        $outcome = 'failed'
    } elseif ($reviewRows.Count -gt 0 -or $failedRows.Count -gt 0 -or $criticalRows.Count -gt 0) {
        $outcome = 'needs_review'
    }

    $criticalCodeText = (($criticalRows | ForEach-Object { [string]$_.错误码 }) -join ';')
    $failureDetails = (($failedRows | Select-Object -First 10 | ForEach-Object {
        $product = if ([string]::IsNullOrWhiteSpace([string]$_.商品文件名)) { [string]$_.货号 } else { [string]$_.商品文件名 }
        $reason = if ([string]::IsNullOrWhiteSpace([string]$_.中文说明)) { [string]$_.错误码 } else { [string]$_.中文说明 }
        "$product：$reason"
    }) -join '；')
    return [pscustomobject]@{
        Outcome = $outcome
        TotalCount = $rows.Count
        SuccessCount = $successRows.Count
        ReviewCount = $reviewRows.Count
        FailureCount = $failedRows.Count
        CriticalCount = $criticalRows.Count
        ExportedCount = $exportedCount
        CriticalCodes = $criticalCodeText
        FailureDetails = $failureDetails
        SummaryText = "导出 JPG：$exportedCount 张；成功：$($successRows.Count) 条；需复核：$($reviewRows.Count) 条；失败：$($failedRows.Count) 条；总计：$($rows.Count) 条。"
    }
}

function ConvertTo-JsStringLiteral {
    param([string]$Value)
    if ($null -eq $Value) {
        $Value = ''
    }
    $escaped = $Value.Replace('\', '\\').Replace('"', '\"').Replace("`r", '\r').Replace("`n", '\n')
    return '"' + $escaped + '"'
}

function Select-File {
    param(
        [string]$Title,
        [string]$Filter
    )
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = $Title
    $dialog.Filter = $Filter
    $dialog.Multiselect = $false
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.FileName
    }
    return $null
}

function Select-Folder {
    param([string]$Description)
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = $Description
    $dialog.ShowNewFolderButton = $true
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.SelectedPath
    }
    return $null
}

function Read-UserSettings {
    if (-not (Test-Path -LiteralPath $settingsPath)) {
        Add-Log "运行偏好文件不存在，将使用空默认值：$settingsPath"
        return [pscustomobject]@{}
    }
    try {
        $text = Get-Content -LiteralPath $settingsPath -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($text)) {
            return [pscustomobject]@{}
        }
        $settings = $text | ConvertFrom-Json
        Add-Log "已读取运行偏好文件：$settingsPath"
        return $settings
    } catch {
        Add-Log "读取运行偏好文件失败，将忽略并继续：$($_.Exception.Message)"
        return [pscustomobject]@{}
    }
}

function Get-SettingText {
    param(
        [object]$Settings,
        [string]$Name
    )
    if ($Settings -and ($Settings.PSObject.Properties.Name -contains $Name)) {
        return [string]$Settings.$Name
    }
    return ''
}

function Save-UserSettings {
    param(
        [string]$ExcelPath,
        [string]$PsdPath,
        [string]$OutputRoot,
        [string]$SheetName,
        [string]$ProductName
    )
    $payload = [pscustomobject]@{
        excelPath = $ExcelPath
        psdPath = $PsdPath
        outputRoot = $OutputRoot
        sheetName = $SheetName
        productName = $ProductName
        updatedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    }
    try {
        $json = $payload | ConvertTo-Json -Depth 4
        Write-Utf8Bom -Path $settingsPath -Content ($json + [Environment]::NewLine)
        $script:Settings = $payload
        Add-Log "已保存运行偏好：$settingsPath"
    } catch {
        Add-Log "保存运行偏好失败，不影响本次任务继续：$($_.Exception.Message)"
    }
}

function New-RunProgressWindow {
    if ($NoUi) {
        return
    }
    $form = New-Object System.Windows.Forms.Form
    $form.Text = '电商主图批量套版正在执行'
    $form.StartPosition = 'CenterScreen'
    $form.Size = New-Object System.Drawing.Size(720, 210)
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.ControlBox = $false
    $form.MinimizeBox = $true
    $form.MaximizeBox = $false
    $form.TopMost = $false

    $title = New-Object System.Windows.Forms.Label
    $title.Text = '任务正在执行。可以最小化此窗口，请勿关闭 Photoshop。'
    $title.AutoSize = $true
    $title.Location = New-Object System.Drawing.Point(18, 18)
    $form.Controls.Add($title)

    $stage = New-Object System.Windows.Forms.Label
    $stage.Text = '正在准备任务...'
    $stage.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 11, [System.Drawing.FontStyle]::Bold)
    $stage.AutoSize = $true
    $stage.Location = New-Object System.Drawing.Point(18, 55)
    $form.Controls.Add($stage)

    $detail = New-Object System.Windows.Forms.Label
    $detail.Text = '正在初始化。'
    $detail.AutoEllipsis = $true
    $detail.Size = New-Object System.Drawing.Size(670, 42)
    $detail.Location = New-Object System.Drawing.Point(18, 86)
    $form.Controls.Add($detail)

    $bar = New-Object System.Windows.Forms.ProgressBar
    $bar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
    $bar.MarqueeAnimationSpeed = 30
    $bar.Size = New-Object System.Drawing.Size(670, 18)
    $bar.Location = New-Object System.Drawing.Point(18, 140)
    $form.Controls.Add($bar)

    $script:ProgressForm = $form
    $script:ProgressStageLabel = $stage
    $script:ProgressDetailLabel = $detail
    $form.Show()
    [System.Windows.Forms.Application]::DoEvents()
}

function Set-RunProgress {
    param(
        [string]$Stage,
        [string]$Detail
    )
    $script:Stage = $Stage
    Add-Log "进度：$Stage - $Detail"
    Write-EntryStatus -Status 'running' -Message $Detail
    if ($script:ProgressForm -and -not $script:ProgressForm.IsDisposed) {
        $script:ProgressStageLabel.Text = $Stage
        $script:ProgressDetailLabel.Text = $Detail
        $script:ProgressForm.Refresh()
        [System.Windows.Forms.Application]::DoEvents()
    }
}

function Close-RunProgressWindow {
    if ($script:ProgressForm -and -not $script:ProgressForm.IsDisposed) {
        try { $script:ProgressForm.Close() } catch {}
    }
    $script:ProgressForm = $null
}

function Select-EditableFolderPath {
    param(
        [string]$Title,
        [string]$Description,
        [string]$InitialPath,
        [string]$BrowseDescription
    )
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.StartPosition = 'CenterScreen'
    $form.Size = New-Object System.Drawing.Size(720, 240)
    $form.MinimizeBox = $false
    $form.MaximizeBox = $false

    $lastText = if ([string]::IsNullOrWhiteSpace($InitialPath)) { '无' } else { $InitialPath }
    $label = New-Object System.Windows.Forms.Label
    $label.Text = "$Description`r`n可直接输入已授权 SMB/UNC 路径，例如：\\10.1.212.3\共享目录\项目`r`n上次保存：$lastText"
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(16, 16)
    $form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(16, 92)
    $textBox.Size = New-Object System.Drawing.Size(560, 28)
    $textBox.Text = $InitialPath
    $form.Controls.Add($textBox)

    $browseButton = New-Object System.Windows.Forms.Button
    $browseButton.Text = '浏览本地目录'
    $browseButton.Location = New-Object System.Drawing.Point(590, 90)
    $browseButton.Size = New-Object System.Drawing.Size(100, 28)
    $browseButton.Add_Click({
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = $BrowseDescription
        $dialog.ShowNewFolderButton = $true
        if (-not [string]::IsNullOrWhiteSpace($textBox.Text) -and (Test-Path -LiteralPath $textBox.Text -PathType Container)) {
            $dialog.SelectedPath = $textBox.Text
        }
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $textBox.Text = $dialog.SelectedPath
        }
    })
    $form.Controls.Add($browseButton)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = '确定'
    $okButton.Location = New-Object System.Drawing.Point(482, 150)
    $okButton.Add_Click({
        if ([string]::IsNullOrWhiteSpace($textBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show('路径不能为空。', $Title) | Out-Null
            return
        }
        $form.Tag = $textBox.Text.Trim()
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })
    $form.Controls.Add($okButton)
    $form.AcceptButton = $okButton

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = '取消'
    $cancelButton.Location = New-Object System.Drawing.Point(592, 150)
    $cancelButton.Add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.Close()
    })
    $form.Controls.Add($cancelButton)
    $form.CancelButton = $cancelButton

    if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return [string]$form.Tag
    }
    return $null
}

function Assert-ReadableDirectory {
    param(
        [string]$Path,
        [string]$Purpose
    )
    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw "$Purpose 路径为空。"
    }
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "$Purpose 路径不存在：$Path。请确认路径存在；如果是 SMB/UNC 地址，请确认已在 Windows 中完成访问授权。"
    }

    try {
        $item = Get-Item -LiteralPath $Path -ErrorAction Stop
    } catch {
        throw "$Purpose 路径不可访问或无权限：$Path。请确认路径存在；如果是 SMB/UNC 地址，请确认已在 Windows 中完成访问授权。原始错误：$($_.Exception.Message)"
    }
    if (-not $item.PSIsContainer) {
        throw "$Purpose 路径不是文件夹：$Path。"
    }

    try {
        $enumerator = [System.IO.Directory]::EnumerateFileSystemEntries($item.FullName).GetEnumerator()
        try {
            $null = $enumerator.MoveNext()
        } finally {
            if ($enumerator -is [System.IDisposable]) {
                $enumerator.Dispose()
            }
        }
    } catch {
        throw "$Purpose 路径不可读取或无权限：$Path。请确认路径可访问；如果是 SMB/UNC 地址，请确认已在 Windows 中完成访问授权。原始错误：$($_.Exception.Message)"
    }
}

function Assert-WritableDirectory {
    param(
        [string]$Path,
        [string]$Purpose
    )
    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw "$Purpose 路径为空。"
    }
    if (-not (Test-Path -LiteralPath $Path)) {
        try {
            New-Item -Path $Path -ItemType Directory -Force -ErrorAction Stop | Out-Null
            Add-Log "已创建$Purpose：$Path"
        } catch {
            throw "$Purpose 路径不可创建：$Path。请换一个可写目录，或确认 SMB/UNC 地址已有写入权限。原始错误：$($_.Exception.Message)"
        }
    }

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "$Purpose 路径不存在：$Path。请确认路径存在；如果是 SMB/UNC 地址，请确认已在 Windows 中完成访问授权。"
    }

    try {
        $item = Get-Item -LiteralPath $Path -ErrorAction Stop
    } catch {
        throw "$Purpose 路径不可访问或无权限：$Path。请确认路径存在；如果是 SMB/UNC 地址，请确认已在 Windows 中完成访问授权。原始错误：$($_.Exception.Message)"
    }
    if (-not $item.PSIsContainer) {
        throw "$Purpose 路径不是文件夹：$Path。"
    }

    try {
        $enumerator = [System.IO.Directory]::EnumerateFileSystemEntries($item.FullName).GetEnumerator()
        try {
            $null = $enumerator.MoveNext()
        } finally {
            if ($enumerator -is [System.IDisposable]) {
                $enumerator.Dispose()
            }
        }
    } catch {
        throw "$Purpose 路径不可读取或无权限：$Path。请确认路径可访问；如果是 SMB/UNC 地址，请确认已在 Windows 中完成访问授权。原始错误：$($_.Exception.Message)"
    }

    $testPath = Join-Path $Path ("_l0_write_test_{0}.tmp" -f ([guid]::NewGuid().ToString('N')))
    try {
        Set-Content -LiteralPath $testPath -Value 'L0 write test' -Encoding UTF8 -ErrorAction Stop
        Remove-Item -LiteralPath $testPath -Force -ErrorAction Stop
    } catch {
        throw "$Purpose 路径不可写或无权限：$Path。请换一个可写目录，或确认 SMB/UNC 地址已有写入权限。原始错误：$($_.Exception.Message)"
    }
}

function Select-Sheet {
    param(
        [string[]]$Sheets,
        [string]$WorkbookName
    )
    $form = New-Object System.Windows.Forms.Form
    $form.Text = '选择数据工作表'
    $form.StartPosition = 'CenterScreen'
    $form.Size = New-Object System.Drawing.Size(520, 420)
    $form.MinimizeBox = $false
    $form.MaximizeBox = $false

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "表格：$WorkbookName`r`n请选择要使用的数据工作表："
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(16, 16)
    $form.Controls.Add($label)

    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Location = New-Object System.Drawing.Point(16, 64)
    $listBox.Size = New-Object System.Drawing.Size(470, 240)
    foreach ($sheet in $Sheets) {
        [void]$listBox.Items.Add($sheet)
    }
    if ($listBox.Items.Count -gt 0) {
        $listBox.SelectedIndex = 0
    }
    $form.Controls.Add($listBox)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = '确定'
    $okButton.Location = New-Object System.Drawing.Point(300, 325)
    $okButton.Add_Click({
        if ($listBox.SelectedItem) {
            $form.Tag = [string]$listBox.SelectedItem
            $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $form.Close()
        }
    })
    $form.Controls.Add($okButton)
    $form.AcceptButton = $okButton

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = '取消'
    $cancelButton.Location = New-Object System.Drawing.Point(400, 325)
    $cancelButton.Add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.Close()
    })
    $form.Controls.Add($cancelButton)
    $form.CancelButton = $cancelButton

    if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return [string]$form.Tag
    }
    return $null
}

function Select-ProductTask {
    param(
        [string[]]$Products,
        [string]$SheetName
    )
    $form = New-Object System.Windows.Forms.Form
    $form.Text = '选择要生成的商品'
    $form.StartPosition = 'CenterScreen'
    $form.Size = New-Object System.Drawing.Size(640, 460)
    $form.MinimizeBox = $false
    $form.MaximizeBox = $false

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "数据工作表：$SheetName`r`n请选择【全部商品】，或只选择部分商品："
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(16, 16)
    $form.Controls.Add($label)

    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Location = New-Object System.Drawing.Point(16, 64)
    $listBox.Size = New-Object System.Drawing.Size(590, 280)
    [void]$listBox.Items.Add('全部商品')
    foreach ($product in $Products) {
        [void]$listBox.Items.Add($product)
    }
    $listBox.SelectedIndex = 0
    $form.Controls.Add($listBox)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = '确定'
    $okButton.Location = New-Object System.Drawing.Point(420, 370)
    $okButton.Add_Click({
        if ($listBox.SelectedItem) {
            $selected = [string]$listBox.SelectedItem
            if ($selected -eq '全部商品') {
                $form.Tag = ''
            } else {
                $form.Tag = $selected
            }
            $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $form.Close()
        }
    })
    $form.Controls.Add($okButton)
    $form.AcceptButton = $okButton

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = '取消'
    $cancelButton.Location = New-Object System.Drawing.Point(520, 370)
    $cancelButton.Add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.Close()
    })
    $form.Controls.Add($cancelButton)
    $form.CancelButton = $cancelButton

    if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return [string]$form.Tag
    }
    return $null
}

function Move-FormIntoVisibleWorkingArea {
    param([object]$Form)
    if (-not $Form) { return }

    $workingArea = [System.Windows.Forms.Screen]::FromControl($Form).WorkingArea
    $maxLeft = [Math]::Max($workingArea.Left, $workingArea.Right - $Form.Width)
    $maxTop = [Math]::Max($workingArea.Top, $workingArea.Bottom - $Form.Height)
    $left = [Math]::Min([Math]::Max($Form.Left, $workingArea.Left), $maxLeft)
    $top = [Math]::Min([Math]::Max($Form.Top, $workingArea.Top), $maxTop)

    if ($left -ne $Form.Left -or $top -ne $Form.Top) {
        $Form.Location = New-Object System.Drawing.Point($left, $top)
    }
}

function Select-TaskSettings {
    param(
        [object]$Python,
        [string]$InitialExcelPath,
        [string]$InitialPsdPath,
        [string]$InitialOutputRoot,
        [string]$InitialSheetName,
        [string]$InitialProductName
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = '电商主图套版工具 1.0'
    $form.StartPosition = 'CenterScreen'
    $form.Size = New-Object System.Drawing.Size(860, 610)
    $form.MinimumSize = New-Object System.Drawing.Size(860, 610)
    $form.MinimizeBox = $false
    $form.MaximizeBox = $false
    $form.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)
    $form.Add_Shown({ Move-FormIntoVisibleWorkingArea -Form $form })

    $title = New-Object System.Windows.Forms.Label
    $title.Text = '请先启动并登录 Photoshop，进入首页后再选择商品表格和 PSD 模板。'
    $title.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 11, [System.Drawing.FontStyle]::Bold)
    $title.AutoSize = $true
    $title.Location = New-Object System.Drawing.Point(16, 16)
    $form.Controls.Add($title)

    $excelLabel = New-Object System.Windows.Forms.Label
    $excelLabel.Text = '1  商品信息表格'
    $excelLabel.AutoSize = $true
    $excelLabel.Location = New-Object System.Drawing.Point(16, 55)
    $form.Controls.Add($excelLabel)

    $excelBox = New-Object System.Windows.Forms.TextBox
    $excelBox.Location = New-Object System.Drawing.Point(125, 52)
    $excelBox.Size = New-Object System.Drawing.Size(585, 24)
    $excelBox.Text = $InitialExcelPath
    $form.Controls.Add($excelBox)

    $excelBrowse = New-Object System.Windows.Forms.Button
    $excelBrowse.Text = '浏览...'
    $excelBrowse.Location = New-Object System.Drawing.Point(725, 50)
    $excelBrowse.Size = New-Object System.Drawing.Size(88, 28)
    $form.Controls.Add($excelBrowse)

    $psdLabel = New-Object System.Windows.Forms.Label
    $psdLabel.Text = '2  PSD 模板'
    $psdLabel.AutoSize = $true
    $psdLabel.Location = New-Object System.Drawing.Point(16, 93)
    $form.Controls.Add($psdLabel)

    $psdBox = New-Object System.Windows.Forms.TextBox
    $psdBox.Location = New-Object System.Drawing.Point(125, 90)
    $psdBox.Size = New-Object System.Drawing.Size(585, 24)
    $psdBox.Text = $InitialPsdPath
    $form.Controls.Add($psdBox)

    $psdBrowse = New-Object System.Windows.Forms.Button
    $psdBrowse.Text = '浏览...'
    $psdBrowse.Location = New-Object System.Drawing.Point(725, 88)
    $psdBrowse.Size = New-Object System.Drawing.Size(88, 28)
    $form.Controls.Add($psdBrowse)

    $outputLabel = New-Object System.Windows.Forms.Label
    $outputLabel.Text = '成品保存到'
    $outputLabel.AutoSize = $true
    $outputLabel.Location = New-Object System.Drawing.Point(16, 131)
    $form.Controls.Add($outputLabel)

    $outputBox = New-Object System.Windows.Forms.TextBox
    $outputBox.Location = New-Object System.Drawing.Point(125, 128)
    $outputBox.Size = New-Object System.Drawing.Size(480, 24)
    $outputBox.Text = $InitialOutputRoot
    $form.Controls.Add($outputBox)

    $outputBrowse = New-Object System.Windows.Forms.Button
    $outputBrowse.Text = '更改...'
    $outputBrowse.Location = New-Object System.Drawing.Point(620, 126)
    $outputBrowse.Size = New-Object System.Drawing.Size(88, 28)
    $form.Controls.Add($outputBrowse)

    $openOutputButton = New-Object System.Windows.Forms.Button
    $openOutputButton.Text = '打开位置'
    $openOutputButton.Location = New-Object System.Drawing.Point(725, 126)
    $openOutputButton.Size = New-Object System.Drawing.Size(88, 28)
    $form.Controls.Add($openOutputButton)

    $outputHint = New-Object System.Windows.Forms.Label
    $outputHint.Text = '每次任务会自动新建文件夹，不会覆盖以前的成品。'
    $outputHint.AutoSize = $true
    $outputHint.ForeColor = [System.Drawing.Color]::DimGray
    $outputHint.Location = New-Object System.Drawing.Point(125, 158)
    $form.Controls.Add($outputHint)

    $sheetLabel = New-Object System.Windows.Forms.Label
    $sheetLabel.Text = '3  数据工作表'
    $sheetLabel.AutoSize = $true
    $sheetLabel.Location = New-Object System.Drawing.Point(16, 195)
    $form.Controls.Add($sheetLabel)

    $sheetCombo = New-Object System.Windows.Forms.ComboBox
    $sheetCombo.Location = New-Object System.Drawing.Point(125, 192)
    $sheetCombo.Size = New-Object System.Drawing.Size(300, 26)
    $sheetCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $form.Controls.Add($sheetCombo)

    $reloadButton = New-Object System.Windows.Forms.Button
    $reloadButton.Text = '重新读取'
    $reloadButton.Location = New-Object System.Drawing.Point(440, 190)
    $reloadButton.Size = New-Object System.Drawing.Size(96, 28)
    $form.Controls.Add($reloadButton)

    $productLabel = New-Object System.Windows.Forms.Label
    $productLabel.Text = '4  商品范围'
    $productLabel.AutoSize = $true
    $productLabel.Location = New-Object System.Drawing.Point(16, 243)
    $form.Controls.Add($productLabel)

    $allProductsRadio = New-Object System.Windows.Forms.RadioButton
    $allProductsRadio.Text = '全部商品'
    $allProductsRadio.AutoSize = $true
    $allProductsRadio.Location = New-Object System.Drawing.Point(125, 241)
    $form.Controls.Add($allProductsRadio)

    $singleProductRadio = New-Object System.Windows.Forms.RadioButton
    $singleProductRadio.Text = '只生成勾选的商品'
    $singleProductRadio.AutoSize = $true
    $singleProductRadio.Location = New-Object System.Drawing.Point(225, 241)
    $form.Controls.Add($singleProductRadio)

    $productList = New-Object System.Windows.Forms.CheckedListBox
    $productList.Location = New-Object System.Drawing.Point(125, 273)
    $productList.Size = New-Object System.Drawing.Size(688, 166)
    $productList.CheckOnClick = $true
    $productList.Enabled = $false
    $form.Controls.Add($productList)

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Text = '请先选择商品信息表格。'
    $statusLabel.AutoEllipsis = $true
    $statusLabel.Size = New-Object System.Drawing.Size(797, 42)
    $statusLabel.Location = New-Object System.Drawing.Point(16, 454)
    $form.Controls.Add($statusLabel)

    $runButton = New-Object System.Windows.Forms.Button
    $runButton.Text = '开始生成'
    $runButton.Location = New-Object System.Drawing.Point(570, 510)
    $runButton.Size = New-Object System.Drawing.Size(138, 36)
    $form.Controls.Add($runButton)
    $form.AcceptButton = $runButton

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = '取消'
    $cancelButton.Location = New-Object System.Drawing.Point(725, 510)
    $cancelButton.Size = New-Object System.Drawing.Size(88, 36)
    $form.Controls.Add($cancelButton)
    $form.CancelButton = $cancelButton

    $script:LoadedTaskSettingsExcelPath = ''

    $clearSheetAndProducts = {
        $sheetCombo.Items.Clear()
        $productList.Items.Clear()
        $productList.Enabled = $false
        $singleProductRadio.Enabled = $false
        $allProductsRadio.Checked = $true
        $statusLabel.Text = '表格已更换，请重新读取工作表。'
    }

    $setBusy = {
        param([bool]$Busy)
        $excelBrowse.Enabled = -not $Busy
        $psdBrowse.Enabled = -not $Busy
        $outputBrowse.Enabled = -not $Busy
        $openOutputButton.Enabled = -not $Busy
        $reloadButton.Enabled = -not $Busy
        $runButton.Enabled = -not $Busy
        [System.Windows.Forms.Application]::DoEvents()
    }

    $loadProducts = {
        param([string]$PreferredProduct)
        $productList.Items.Clear()
        $productList.Enabled = $false
        $singleProductRadio.Enabled = $false
        $allProductsRadio.Checked = $true
        if (-not $sheetCombo.SelectedItem) {
            return $false
        }

        $productLines = @(Invoke-Python -Python $Python -Arguments @($sheetScript, '--products', $excelBox.Text.Trim(), [string]$sheetCombo.SelectedItem) 2>&1)
        if ($LASTEXITCODE -ne 0) {
            $statusLabel.Text = "读取商品任务列表失败：$($productLines -join '；')"
            return $false
        }

        foreach ($product in @($productLines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
            [void]$productList.Items.Add([string]$product)
        }
        if ($productList.Items.Count -gt 0) {
            $singleProductRadio.Enabled = $true
            if (-not [string]::IsNullOrWhiteSpace($PreferredProduct) -and $PreferredProduct -ne '*' -and $PreferredProduct -ne '全部商品') {
                $preferredProducts = @($PreferredProduct -split '\|' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
                foreach ($preferred in $preferredProducts) {
                    $matchIndex = $productList.Items.IndexOf($preferred)
                    if ($matchIndex -ge 0) {
                        $productList.SetItemChecked($matchIndex, $true)
                    }
                }
                if ($productList.CheckedItems.Count -gt 0) { $singleProductRadio.Checked = $true }
            }
        }
        $productList.Enabled = ($singleProductRadio.Checked -and $singleProductRadio.Enabled)
        $statusLabel.Text = "表格读取完成：$($productList.Items.Count) 个商品。默认生成全部商品。"
        return $true
    }

    $loadSheets = {
        param([string]$PreferredSheet, [string]$PreferredProduct)
        $sheetCombo.Items.Clear()
        $productList.Items.Clear()
        $productList.Enabled = $false
        $singleProductRadio.Enabled = $false
        $allProductsRadio.Checked = $true

        $candidateExcel = $excelBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($candidateExcel)) {
            $statusLabel.Text = 'Excel 变量表不能为空。'
            return $false
        }
        if (-not (Test-Path -LiteralPath $candidateExcel -PathType Leaf)) {
            $statusLabel.Text = "Excel 文件不存在或不可访问：$candidateExcel"
            return $false
        }

        & $setBusy $true
        try {
            $sheetLines = @(Invoke-Python -Python $Python -Arguments @($sheetScript, $candidateExcel) 2>&1)
            if ($LASTEXITCODE -ne 0) {
            $statusLabel.Text = "读取工作表失败：$($sheetLines -join '；')"
                return $false
            }
            foreach ($sheet in @($sheetLines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
                [void]$sheetCombo.Items.Add([string]$sheet)
            }
            if ($sheetCombo.Items.Count -eq 0) {
                $statusLabel.Text = '表格中没有可使用的数据工作表。'
                return $false
            }
            $sheetIndex = -1
            if (-not [string]::IsNullOrWhiteSpace($PreferredSheet)) {
                $sheetIndex = $sheetCombo.Items.IndexOf($PreferredSheet)
            }
            if ($sheetIndex -lt 0) {
                $sheetIndex = 0
            }
            $sheetCombo.SelectedIndex = $sheetIndex
            $statusLabel.Text = "已读取 $($sheetCombo.Items.Count) 个工作表。"
            $script:LoadedTaskSettingsExcelPath = $candidateExcel
            return (& $loadProducts $PreferredProduct)
        } catch {
            $message = $_.Exception.Message
            Add-Log "读取 Excel 设置时发生未处理异常：$message"
            $statusLabel.Text = "读取 Excel 时发生异常：$message"
            return $false
        } finally {
            & $setBusy $false
        }
    }

    $excelBox.Add_TextChanged({
        if ($excelBox.Text.Trim() -ne $script:LoadedTaskSettingsExcelPath) {
            $script:LoadedTaskSettingsExcelPath = ''
            & $clearSheetAndProducts
        }
    })

    $excelBrowse.Add_Click({
        $selected = Select-File -Title '请选择 Excel 变量表（xlsx）' -Filter 'Excel 文件 (*.xlsx)|*.xlsx'
        if ($selected) {
            $excelBox.Text = $selected
            [void](& $loadSheets '' '')
        }
    })

    $psdBrowse.Add_Click({
        $selected = Select-File -Title '请选择 PSD 模板文件' -Filter 'Photoshop PSD (*.psd)|*.psd'
        if ($selected) {
            $psdBox.Text = $selected
        }
    })

    $outputBrowse.Add_Click({
        $selected = Select-Folder -Description '请选择输出保存位置'
        if ($selected) {
            $outputBox.Text = $selected
        }
    })

    $openOutputButton.Add_Click({
        $path = $outputBox.Text.Trim()
        try {
            Assert-WritableDirectory -Path $path -Purpose '成品保存位置'
            Start-Process -FilePath $path
        } catch {
            $statusLabel.Text = "无法打开保存位置：$($_.Exception.Message)"
        }
    })

    $reloadButton.Add_Click({
        [void](& $loadSheets ([string]$sheetCombo.SelectedItem) '')
    })

    $sheetCombo.Add_SelectedIndexChanged({
        [void](& $loadProducts '')
    })

    $allProductsRadio.Add_CheckedChanged({
        $productList.Enabled = ($singleProductRadio.Checked -and $singleProductRadio.Enabled)
    })

    $singleProductRadio.Add_CheckedChanged({
        $productList.Enabled = ($singleProductRadio.Checked -and $singleProductRadio.Enabled)
    })

    $runButton.Add_Click({
        $photoshopCheck = Get-PhotoshopReadiness
        if (-not $photoshopCheck.Ready) {
            $statusLabel.Text = $photoshopCheck.Message
            Add-Log "Photoshop 前置检查未通过：$($photoshopCheck.Code)；$($photoshopCheck.Detail)"
            return
        }
        Add-Log "Photoshop 前置检查通过：版本 $($photoshopCheck.Version)。"
        $candidateExcel = $excelBox.Text.Trim()
        $candidatePsd = $psdBox.Text.Trim()
        $candidateOutput = $outputBox.Text.Trim()
        if (-not (Test-Path -LiteralPath $candidateExcel -PathType Leaf)) {
            $statusLabel.Text = '请重新选择可访问的 Excel 商品信息表格。'
            return
        }
        if ($candidateExcel -ne $script:LoadedTaskSettingsExcelPath) {
            if (-not (& $loadSheets '' '')) { return }
        }
        if (-not (Test-Path -LiteralPath $candidatePsd -PathType Leaf)) {
            $statusLabel.Text = '请重新选择可访问的 PSD 模板。'
            return
        }
        if ([string]::IsNullOrWhiteSpace($candidateOutput)) {
            $candidateOutput = Get-DefaultOutputRoot
            $outputBox.Text = $candidateOutput
        }
        try {
            Assert-WritableDirectory -Path $candidateOutput -Purpose '输出保存位置'
        } catch {
            $statusLabel.Text = "成品保存位置不可用：$($_.Exception.Message)"
            return
        }
        if (-not $sheetCombo.SelectedItem) {
            $statusLabel.Text = '请选择数据工作表。'
            return
        }
        $candidateProduct = ''
        if ($singleProductRadio.Checked) {
            $selectedProducts = @($productList.CheckedItems | ForEach-Object { [string]$_ })
            if ($selectedProducts.Count -eq 0) {
                $statusLabel.Text = '请勾选至少一个商品，或改选“全部商品”。'
                return
            }
            $candidateProduct = $selectedProducts -join '|'
        }
        $form.Tag = [pscustomobject]@{
            ExcelPath = $candidateExcel
            PsdPath = $candidatePsd
            OutputRoot = $candidateOutput
            SheetName = [string]$sheetCombo.SelectedItem
            ProductName = $candidateProduct
        }
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })

    $cancelButton.Add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.Close()
    })

    if (-not [string]::IsNullOrWhiteSpace($InitialExcelPath) -and (Test-Path -LiteralPath $InitialExcelPath -PathType Leaf)) {
        [void](& $loadSheets $InitialSheetName $InitialProductName)
    } elseif (-not [string]::IsNullOrWhiteSpace($InitialExcelPath)) {
        $statusLabel.Text = "上次 Excel 不存在或不可访问，请重新选择：$InitialExcelPath"
    }

    if ([string]::IsNullOrWhiteSpace($InitialProductName) -or $InitialProductName -eq '*' -or $InitialProductName -eq '全部商品') {
        $allProductsRadio.Checked = $true
    }

    if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $form.Tag
    }
    return $null
}

function New-PythonCandidate {
    param(
        [string]$Command,
        [string[]]$PrefixArgs,
        [string]$Label
    )
    return [pscustomobject]@{ Command = $Command; PrefixArgs = $PrefixArgs; Label = $Label }
}

function Get-PrivatePythonCandidate {
    if (Test-Path -LiteralPath $privatePythonExe) {
        Add-Log "工具私有 Python 候选存在：$privatePythonExe"
        return (New-PythonCandidate -Command $privatePythonExe -PrefixArgs @() -Label "工具私有 Python：$privatePythonExe")
    }
    Add-Log "工具私有 Python 候选不存在：$privatePythonExe"
    return $null
}

function Test-IsWindowsAppsAlias {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    return ($Path -like '*\Microsoft\WindowsApps\python*.exe')
}

function Write-SystemPythonProbeLog {
    $pyLauncher = Get-Command py.exe -ErrorAction SilentlyContinue
    if ($pyLauncher) {
        Add-Log "仅记录：发现系统 py.exe -> $($pyLauncher.Source)；本工具不会使用系统 Python。"
    } else {
        Add-Log '仅记录：未发现系统 py.exe；本工具不会使用系统 Python。'
    }

    $python = Get-Command python.exe -ErrorAction SilentlyContinue
    if ($python) {
        if (Test-IsWindowsAppsAlias -Path $python.Source) {
            Add-Log "仅记录：发现 WindowsApps Python 应用商店别名：$($python.Source)；本工具不会使用它。"
        } else {
            Add-Log "仅记录：发现系统 python.exe -> $($python.Source)；本工具不会使用系统 Python。"
        }
    } else {
        Add-Log '仅记录：未发现系统 python.exe；本工具不会使用系统 Python。'
    }
}

function Assert-SupportedWindowsRuntimeHost {
    if (-not [Environment]::Is64BitOperatingSystem) {
        throw '当前系统不是 64 位 Windows。本工具包只支持 Windows 10 x64 或以上。'
    }
    $osVersion = [Environment]::OSVersion.Version
    if ($osVersion.Major -lt 10) {
        throw "当前 Windows 版本 $($osVersion.ToString()) 不支持。本工具包只支持 Windows 10 x64 或以上。"
    }
    Add-Log "Windows 运行宿主预检通过：64 位；版本 $($osVersion.ToString())。"
}

function Invoke-Python {
    param(
        [object]$Python,
        [string[]]$Arguments
    )
    $allArgs = @()
    $allArgs += $Python.PrefixArgs
    $allArgs += $Arguments
    & $Python.Command @allArgs
}

function Test-PrivatePythonRuntime {
    param([object]$Python)
    if (-not $Python) { return $false }
    $probeDiagnosticPath = $runtimeProbeTxt
    $outputLines = @()
    $probeExitCodeText = '未取得'
    try {
        $probe = 'import sys, openpyxl, et_xmlfile; print(sys.version.split()[0]); print(openpyxl.__version__); print(et_xmlfile.__version__)'
        $output = @(Invoke-Python -Python $Python -Arguments @('-c', $probe) 2>&1)
        $outputLines = @($output | ForEach-Object { [string]$_ })
        $probeExitCodeText = [string]$LASTEXITCODE
        foreach ($line in $outputLines) {
            Add-Log "工具私有运行时探针输出：$line"
        }
        if ($LASTEXITCODE -eq 0) {
            $pythonVersion = [string]($outputLines | Select-Object -Index 0)
            $openpyxlVersion = [string]($outputLines | Select-Object -Index 1)
            $etXmlfileVersion = [string]($outputLines | Select-Object -Index 2)
            if ($pythonVersion -eq $requiredPythonVersion -and $openpyxlVersion -eq $requiredOpenpyxlVersion -and $etXmlfileVersion -eq $requiredEtXmlfileVersion) {
                Add-Log "工具私有运行时检测通过：Python $pythonVersion；openpyxl $openpyxlVersion；et-xmlfile $etXmlfileVersion。"
                return $true
            }
            Add-Log "工具私有运行时版本不匹配：需要 Python $requiredPythonVersion / openpyxl $requiredOpenpyxlVersion / et-xmlfile $requiredEtXmlfileVersion；实际 Python $pythonVersion / openpyxl $openpyxlVersion / et-xmlfile $etXmlfileVersion。"
            $probeDiagnosticContent = ((@(
                '工具私有运行时探针失败',
                "时间：$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))",
                "候选：$($Python.Label)",
                "退出码：$LASTEXITCODE",
                "需要版本：Python $requiredPythonVersion / openpyxl $requiredOpenpyxlVersion / et-xmlfile $requiredEtXmlfileVersion",
                "实际版本：Python $pythonVersion / openpyxl $openpyxlVersion / et-xmlfile $etXmlfileVersion",
                '',
                'stdout+stderr：'
            ) + $outputLines) -join [Environment]::NewLine) + [Environment]::NewLine
            Write-Utf8Bom -Path $probeDiagnosticPath -Content $probeDiagnosticContent
            return $false
        }
        Add-Log "工具私有运行时检测失败：$($Python.Label)；退出码 $LASTEXITCODE。完整输出已写入 $probeDiagnosticPath"
        $probeDiagnosticContent = ((@(
            '工具私有运行时探针失败',
            "时间：$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))",
            "候选：$($Python.Label)",
            "退出码：$LASTEXITCODE",
            "需要版本：Python $requiredPythonVersion / openpyxl $requiredOpenpyxlVersion / et-xmlfile $requiredEtXmlfileVersion",
            '',
            'stdout+stderr：'
        ) + $outputLines) -join [Environment]::NewLine) + [Environment]::NewLine
        Write-Utf8Bom -Path $probeDiagnosticPath -Content $probeDiagnosticContent
    } catch {
        Add-Log "工具私有运行时检测异常：$($Python.Label)；$($_.Exception.Message)"
        $probeOutputForDiagnostic = @($outputLines)
        if ($probeOutputForDiagnostic.Count -eq 0) {
            $probeOutputForDiagnostic = @('未取得/空')
        }
        $probeDiagnosticContent = ((@(
            '工具私有运行时探针异常',
            "时间：$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))",
            "候选：$($Python.Label)",
            "退出码：$probeExitCodeText",
            "需要版本：Python $requiredPythonVersion / openpyxl $requiredOpenpyxlVersion / et-xmlfile $requiredEtXmlfileVersion",
            "实际版本：Python 未取得 / openpyxl 未取得 / et-xmlfile 未取得",
            '',
            'stdout+stderr：'
        ) + $probeOutputForDiagnostic + @(
            '',
            '异常信息：',
            [string]$_.Exception.Message
        )) -join [Environment]::NewLine) + [Environment]::NewLine
        Write-Utf8Bom -Path $probeDiagnosticPath -Content $probeDiagnosticContent
    }
    return $false
}

function Enable-EmbeddedPythonSite {
    $pthFiles = @(Get-ChildItem -LiteralPath $privatePythonDir -Filter 'python*._pth' -File -ErrorAction SilentlyContinue)
    foreach ($pth in $pthFiles) {
        $lines = [System.IO.File]::ReadAllLines($pth.FullName, [System.Text.Encoding]::UTF8)
        $changed = $false
        for ($i = 0; $i -lt $lines.Length; $i++) {
            if ($lines[$i] -eq '#import site') {
                $lines[$i] = 'import site'
                $changed = $true
            }
        }
        if ($changed) {
            $encoding = New-Object System.Text.UTF8Encoding($false)
            [System.IO.File]::WriteAllLines($pth.FullName, $lines, $encoding)
            Add-Log "已启用 embedded Python import site：$($pth.FullName)"
        }
    }
}

function Invoke-DownloadFile {
    param(
        [string]$Url,
        [string]$Destination
    )
    Add-Log "下载：$Url"
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    } catch {
    }
    Invoke-WebRequest -Uri $Url -OutFile $Destination -UseBasicParsing -ErrorAction Stop
    if (-not (Test-Path -LiteralPath $Destination)) {
        throw "下载后未找到文件：$Destination"
    }
}

function Install-PrivatePythonRuntime {
    $script:Stage = '自动安装工具私有运行时'
    Add-Log '自动安装工具私有运行时：开始。'
    Assert-SupportedWindowsRuntimeHost

    New-Item -Path $runtimeRoot -ItemType Directory -Force | Out-Null
    if (Test-Path -LiteralPath $runtimeTempDir) {
        Remove-Item -LiteralPath $runtimeTempDir -Recurse -Force
    }
    New-Item -Path $runtimeTempDir -ItemType Directory -Force | Out-Null

    $zipPath = Join-Path $runtimeTempDir 'python-3.11.9-embed-amd64.zip'
    $extractDir = Join-Path $runtimeTempDir 'python_extract'
    $getPipPath = Join-Path $runtimeTempDir 'get-pip.py'
    try {
        Invoke-DownloadFile -Url $pythonEmbedUrl -Destination $zipPath
        New-Item -Path $extractDir -ItemType Directory -Force | Out-Null
        Expand-Archive -LiteralPath $zipPath -DestinationPath $extractDir -Force

        if (Test-Path -LiteralPath $privatePythonDir) {
            Remove-Item -LiteralPath $privatePythonDir -Recurse -Force
        }
        Move-Item -LiteralPath $extractDir -Destination $privatePythonDir
        Enable-EmbeddedPythonSite

        Invoke-DownloadFile -Url $getPipUrl -Destination $getPipPath
        $private = New-PythonCandidate -Command $privatePythonExe -PrefixArgs @() -Label "工具私有 Python：$privatePythonExe"
        $pipOutput = @(Invoke-Python -Python $private -Arguments @($getPipPath, '--no-warn-script-location') 2>&1)
        foreach ($line in $pipOutput) { Add-Log $line }
        if ($LASTEXITCODE -ne 0) {
            throw "pip 安装失败：$($pipOutput -join '；')"
        }

        $openpyxlOutput = @(Invoke-Python -Python $private -Arguments @('-m', 'pip', '--isolated', 'install', '--index-url', 'https://pypi.org/simple', '--no-input', '--disable-pip-version-check', "openpyxl==$requiredOpenpyxlVersion", "et-xmlfile==$requiredEtXmlfileVersion") 2>&1)
        foreach ($line in $openpyxlOutput) { Add-Log $line }
        if ($LASTEXITCODE -ne 0) {
            throw "openpyxl 安装失败：$($openpyxlOutput -join '；')"
        }
        Add-Log '自动安装工具私有运行时：完成。'
    } catch {
        Add-Log "自动安装工具私有运行时失败：$($_.Exception.Message)"
        throw
    } finally {
        try {
            if (Test-Path -LiteralPath $runtimeTempDir) {
                Remove-Item -LiteralPath $runtimeTempDir -Recurse -Force
            }
        } catch {
            Add-Log "清理运行时临时目录失败：$($_.Exception.Message)"
        }
    }
}

function Ensure-PythonRuntime {
    $script:Stage = '预检工具私有运行时'
    Add-Log '开始预检工具包内置 Python/openpyxl/et-xmlfile 运行环境。'
    Assert-SupportedWindowsRuntimeHost
    Write-SystemPythonProbeLog
    $private = Get-PrivatePythonCandidate
    if (Test-PrivatePythonRuntime -Python $private) {
        Add-Log '检测到完整随包运行时，本次不联网、不安装，直接进入文件选择。'
        return $private
    }

    $script:Stage = '自动修复工具私有运行时'
    Add-Log '工具包内置运行时缺失、损坏或版本不匹配，准备联网从官方源修复 embedded Python 与固定版本 openpyxl/et-xmlfile。'
    Install-PrivatePythonRuntime

    $private = Get-PrivatePythonCandidate
    if (Test-PrivatePythonRuntime -Python $private) {
        return $private
    }
    throw "自动修复后工具私有运行时仍未通过固定版本检测。请确认网络可访问 python.org、bootstrap.pypa.io、pypi.org，且当前目录允许写入。"
}

function Get-RegistryValueSafe {
    param(
        [string]$Path,
        [string]$Name
    )
    try {
        if ($Name -eq '(default)') {
            $key = Get-Item -Path $Path -ErrorAction Stop
            return $key.GetValue('')
        }
        $item = Get-ItemProperty -Path $Path -ErrorAction Stop
        return $item.$Name
    } catch {
        return $null
    }
}

function Get-PhotoshopComProgIds {
    $progIds = New-Object System.Collections.Generic.List[string]
    $progIds.Add('Photoshop.Application') | Out-Null
    foreach ($root in @(
        'Registry::HKEY_CLASSES_ROOT',
        'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Classes',
        'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Classes',
        'Registry::HKEY_CURRENT_USER\SOFTWARE\Classes'
    )) {
        try {
            $items = Get-ChildItem -Path $root -ErrorAction Stop |
                Where-Object { $_.PSChildName -like 'Photoshop.Application*' } |
                Select-Object -ExpandProperty PSChildName
            foreach ($item in $items) {
                if ($item -match '^Photoshop\.Application(\.\d+)?$') {
                    $progIds.Add($item) | Out-Null
                }
            }
        } catch {
        }
    }
    return $progIds | Sort-Object -Unique -Descending
}

function Get-PhotoshopReadiness {
    try {
        $currentSessionId = (Get-Process -Id $PID -ErrorAction Stop).SessionId
        $photoshopProcesses = @(
            Get-Process -Name 'Photoshop' -ErrorAction SilentlyContinue |
                Where-Object { $_.SessionId -eq $currentSessionId }
        )
    } catch {
        $photoshopProcesses = @()
    }

    if ($photoshopProcesses.Count -eq 0) {
        return [pscustomobject]@{
            Ready = $false
            Code = 'PS_NOT_RUNNING'
            Version = ''
            Message = '请先启动并登录 Photoshop，进入首页后再点击“开始生成”。'
            Detail = '当前 Windows 会话中未检测到 Photoshop 进程。'
        }
    }

    $photoshopProcess = $photoshopProcesses | Sort-Object StartTime -Descending | Select-Object -First 1
    if (-not $photoshopProcess.Responding) {
        return [pscustomobject]@{
            Ready = $false
            Code = 'PS_NOT_RESPONDING'
            Version = ''
            Message = 'Photoshop 当前无响应。请先关闭并重新启动 Photoshop，进入首页后再试。'
            Detail = "Photoshop 进程无响应；进程 ID $($photoshopProcess.Id)。"
        }
    }
    if ([int64]$photoshopProcess.MainWindowHandle -eq 0) {
        return [pscustomobject]@{
            Ready = $false
            Code = 'PS_NOT_READY'
            Version = ''
            Message = 'Photoshop 正在启动或尚未显示首页。请处理登录、授权或弹窗，进入首页后再试。'
            Detail = "Photoshop 进程存在但没有主窗口；进程 ID $($photoshopProcess.Id)。"
        }
    }

    $version = ''
    try {
        $photoshopPath = [string]$photoshopProcess.Path
        if ([string]::IsNullOrWhiteSpace($photoshopPath)) {
            $photoshopPath = [string]$photoshopProcess.MainModule.FileName
        }
        if (-not [string]::IsNullOrWhiteSpace($photoshopPath)) {
            $fileVersionInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($photoshopPath)
            $version = [string]$fileVersionInfo.ProductVersion
            if ([string]::IsNullOrWhiteSpace($version)) {
                $version = [string]$fileVersionInfo.FileVersion
            }
        }
    } catch {
        $version = ''
    }
    if ([string]::IsNullOrWhiteSpace($version)) {
        return [pscustomobject]@{
            Ready = $false
            Code = 'PS_VERSION_UNAVAILABLE'
            Version = ''
            Message = '无法读取 Photoshop 版本。请确认使用 Windows 正式安装版，并尝试重启或修复 Photoshop。'
            Detail = "Photoshop 进程存在但无法读取安装版本；进程 ID $($photoshopProcess.Id)。"
        }
    }

    return [pscustomobject]@{
        Ready = $true
        Code = 'PS_READY'
        Version = $version
        Message = "Photoshop $version 已就绪。"
        Detail = "进程 ID $($photoshopProcess.Id)；主窗口句柄 $($photoshopProcess.MainWindowHandle)。"
    }
}

function Start-Photoshop {
    $readiness = Get-PhotoshopReadiness
    if (-not $readiness.Ready) {
        throw "$($readiness.Message) 错误码：$($readiness.Code)。"
    }
    foreach ($progId in Get-PhotoshopComProgIds) {
        try {
            $application = New-Object -ComObject $progId -ErrorAction Stop
            $application.Visible = $true
            Add-Log "Photoshop COM 启动成功：$progId；版本 $($application.Version)"
            return $application
        } catch {
            Add-Log "Photoshop COM 启动失败：$progId；$($_.Exception.Message)"
        }
    }
    throw "Photoshop 自动化接口不可用（检测到版本 $($readiness.Version)）。可能是安装注册异常或当前版本不兼容；请重启或修复 Photoshop 后重试。"
}

function Invoke-PhotoshopJavaScript {
    param(
        [object]$Application,
        [string]$ScriptText
    )
    try {
        return $Application.DoJavaScript($ScriptText, @(), 2)
    } catch {
        $firstError = $_.Exception.Message
    }
    try {
        return $Application.DoJavaScript($ScriptText, @())
    } catch {
        throw "DoJavaScript 调用失败：$firstError / $($_.Exception.Message)"
    }
}

function Close-OpenedPhotoshopDocument {
    if ($script:OpenedDocument) {
        try {
            $script:OpenedDocument.Close(2)
            Add-Log '已关闭本次打开的 PSD 模板，未保存改动。'
        } catch {
            Add-Log "关闭本次打开的 PSD 模板失败：$($_.Exception.Message)"
        } finally {
            $script:OpenedDocument = $null
        }
    }
}

function ConvertFrom-TemplatePreparationResult {
    param([object]$RawResult)
    $text = [string]$RawResult
    $parts = $text -split '\|', 3
    return [pscustomobject]@{
        Status = if ($parts.Count -gt 0) { $parts[0] } else { 'ERROR' }
        Message = if ($parts.Count -gt 1) { $parts[1] } else { '未取得 PSD 检测结果。' }
        TemplatePath = if ($parts.Count -gt 2) { $parts[2] } else { '' }
    }
}

function Invoke-TemplatePreparationCheck {
    param(
        [object]$Application,
        [string]$TemplatePath,
        [string]$Mode
    )
    if (-not (Test-Path -LiteralPath $templatePrepareScript -PathType Leaf)) {
        throw "缺少 PSD 模板检测脚本：$templatePrepareScript"
    }
    $script:OpenedDocument = $Application.Open($TemplatePath)
    try {
        $templateText = [System.IO.File]::ReadAllText($templatePrepareScript, [System.Text.Encoding]::UTF8)
        $prefix = "`$.global.__TEMPLATE_PREP_INPUTS__ = { mode: '" + $Mode + "' };"
        $scriptText = $prefix + "`r`n" + $templateText
        $rawResult = Invoke-PhotoshopJavaScript -Application $Application -ScriptText $scriptText
        $result = ConvertFrom-TemplatePreparationResult -RawResult $rawResult
        Add-Log "PSD 模板检测结果：$($result.Status)；$($result.Message)"
        return $result
    } finally {
        Close-OpenedPhotoshopDocument
    }
}

function Resolve-TemplateForTask {
    param(
        [object]$Application,
        [string]$TemplatePath
    )
    $check = Invoke-TemplatePreparationCheck -Application $Application -TemplatePath $TemplatePath -Mode 'check'
    if ($check.Status -eq 'READY') {
        return $TemplatePath
    }
    if ($check.Status -eq 'AMBIGUOUS') {
        throw "PSD 模板未完成智能化改造，且无法安全自动判断：$($check.Message)"
    }
    if ($check.Status -ne 'NEEDS_PREP') {
        throw "PSD 模板检测失败：$($check.Message)"
    }
    if ($NoUi) {
        throw "PSD 模板未完成智能化改造：$($check.Message)"
    }

    $confirmText = "当前主图模板还不能直接套图：`r`n$($check.Message)`r`n`r`n是否自动生成一份可套图的模板副本并继续？`r`n原始模板不会被修改。"
    $choice = [System.Windows.Forms.MessageBox]::Show(
        $confirmText,
        '主图模板检查',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($choice -ne [System.Windows.Forms.DialogResult]::Yes) {
        throw '已取消 PSD 自动智能化改造。本次未开始生成。'
    }

    Set-RunProgress -Stage '智能化改造 PSD 模板' -Detail '正在生成模板副本并转换商品图智能对象，原始 PSD 不会被修改。'
    $prepared = Invoke-TemplatePreparationCheck -Application $Application -TemplatePath $TemplatePath -Mode 'prepare'
    if ($prepared.Status -ne 'PREPARED' -or [string]::IsNullOrWhiteSpace($prepared.TemplatePath) -or -not (Test-Path -LiteralPath $prepared.TemplatePath -PathType Leaf)) {
        throw "PSD 自动智能化改造未完成：$($prepared.Message)"
    }
    Add-Log "PSD 模板已自动改造为副本：$($prepared.TemplatePath)"
    return $prepared.TemplatePath
}

function Get-TaskElapsedText {
    param([datetime]$StartedAt)
    if (-not $StartedAt) {
        return '耗时未记录'
    }
    $seconds = [math]::Max(1, [int][math]::Round(((Get-Date) - $StartedAt).TotalSeconds, 0, [System.MidpointRounding]::AwayFromZero))
    if ($seconds -lt 60) {
        return "本次任务耗时：$seconds 秒"
    }
    $minutes = [math]::Floor($seconds / 60)
    $remainingSeconds = $seconds % 60
    return "本次任务耗时：$minutes 分 $remainingSeconds 秒"
}

function Show-TaskCompletionDialog {
    param(
        [string]$Title,
        [string]$Message,
        [string]$JpgOutputDir
    )
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.StartPosition = 'CenterScreen'
    $form.Size = New-Object System.Drawing.Size(700, 310)
    $form.MinimizeBox = $false
    $form.MaximizeBox = $false

    $content = New-Object System.Windows.Forms.TextBox
    $content.Multiline = $true
    $content.ReadOnly = $true
    $content.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $content.BackColor = $form.BackColor
    $content.WordWrap = $true
    $content.Text = $Message
    $content.Location = New-Object System.Drawing.Point(20, 22)
    $content.Size = New-Object System.Drawing.Size(645, 180)
    $form.Controls.Add($content)

    $viewButton = New-Object System.Windows.Forms.Button
    $viewButton.Text = '立即查看'
    $viewButton.Location = New-Object System.Drawing.Point(438, 224)
    $viewButton.Size = New-Object System.Drawing.Size(112, 32)
    $viewButton.Add_Click({
        Start-Process -FilePath $JpgOutputDir
        $form.Close()
    })
    $form.Controls.Add($viewButton)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = '确定'
    $okButton.Location = New-Object System.Drawing.Point(563, 224)
    $okButton.Size = New-Object System.Drawing.Size(100, 32)
    $okButton.Add_Click({ $form.Close() })
    $form.Controls.Add($okButton)
    $form.AcceptButton = $okButton
    $form.CancelButton = $okButton
    [void]$form.ShowDialog()
}

try {
    $script:Stage = '固定输出初始化'
    New-Item -Path $diagnosticDir -ItemType Directory -Force | Out-Null
    Set-Content -LiteralPath $ps1Marker -Value ('PS1已启动：' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')) -Encoding UTF8
    Write-EntryStatus -Status 'running' -Message 'PowerShell 入口已启动。'
    Add-Log 'L0 电商主图批量套版启动。'
    Clear-StaleDiagnostics

    $script:Stage = '加载 Windows 弹窗组件'
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    Add-Log 'Windows 弹窗组件加载完成。'

    Set-RunProgress -Stage '预检运行环境' -Detail '正在检查工具自带运行环境。首次修复运行环境时需要联网。'
    $python = Ensure-PythonRuntime
    $script:Settings = Read-UserSettings

    if ($NoUi) {
        Set-RunProgress -Stage '校验命令行参数' -Detail '正在校验命令行传入的 Excel、PSD 和输出目录。'
        if ([string]::IsNullOrWhiteSpace($ExcelPath)) { throw '命令行模式缺少 -ExcelPath。' }
        if (-not (Test-Path -LiteralPath $ExcelPath -PathType Leaf)) { throw "Excel 文件不存在：$ExcelPath" }
        $excelPath = $ExcelPath
        Add-Log "命令行指定 Excel：$excelPath"

        if ([string]::IsNullOrWhiteSpace($PsdPath)) { throw '命令行模式缺少 -PsdPath。' }
        if (-not (Test-Path -LiteralPath $PsdPath -PathType Leaf)) { throw "PSD 模板不存在：$PsdPath" }
        $psdPath = $PsdPath
        Add-Log "命令行指定 PSD：$psdPath"

        if ([string]::IsNullOrWhiteSpace($OutputRoot)) { throw '命令行模式缺少 -OutputRoot。' }
        $outputRoot = $OutputRoot
        $sheetName = $SheetName
        $selectedProduct = $ProductName
        Add-Log "命令行指定输出目录：$outputRoot"
        Assert-WritableDirectory -Path $outputRoot -Purpose '输出保存位置'
    } else {
        Set-RunProgress -Stage '确认任务设置' -Detail '请在弹出的设置窗口确认 Excel、PSD、输出位置、Sheet 和商品范围。'
        $initialExcelPath = if (-not [string]::IsNullOrWhiteSpace($ExcelPath)) { $ExcelPath } else { Get-SettingText -Settings $script:Settings -Name 'excelPath' }
        $initialPsdPath = if (-not [string]::IsNullOrWhiteSpace($PsdPath)) { $PsdPath } else { Get-SettingText -Settings $script:Settings -Name 'psdPath' }
        $initialOutputRoot = if (-not [string]::IsNullOrWhiteSpace($OutputRoot)) { $OutputRoot } else { Get-SettingText -Settings $script:Settings -Name 'outputRoot' }
        if ([string]::IsNullOrWhiteSpace($initialOutputRoot)) { $initialOutputRoot = Get-DefaultOutputRoot }
        $initialSheetName = if (-not [string]::IsNullOrWhiteSpace($SheetName)) { $SheetName } else { Get-SettingText -Settings $script:Settings -Name 'sheetName' }
        $initialProductName = if (-not [string]::IsNullOrWhiteSpace($ProductName)) { $ProductName } else { Get-SettingText -Settings $script:Settings -Name 'productName' }
        $taskSettings = Select-TaskSettings -Python $python -InitialExcelPath $initialExcelPath -InitialPsdPath $initialPsdPath -InitialOutputRoot $initialOutputRoot -InitialSheetName $initialSheetName -InitialProductName $initialProductName
        if (-not $taskSettings) {
            Write-EntryStatus -Status 'cancelled' -Message '用户取消了本次任务。'
            Write-TaskHistory -Status '已取消' -Message '用户在任务设置窗口取消。'
            exit 0
        }
        $excelPath = $taskSettings.ExcelPath
        $psdPath = $taskSettings.PsdPath
        $outputRoot = $taskSettings.OutputRoot
        $sheetName = $taskSettings.SheetName
        $selectedProduct = $taskSettings.ProductName
        Add-Log "窗口确认 Excel：$excelPath"
        Add-Log "窗口确认 PSD：$psdPath"
        Add-Log "窗口确认输出目录：$outputRoot"
        Add-Log "窗口确认 Sheet：$sheetName"
        Add-Log "窗口确认商品任务：$(if ($selectedProduct) { $selectedProduct } else { '全部商品' })"
    }

    $script:CurrentExcelPath = $excelPath
    $script:CurrentPsdPath = $psdPath
    $script:CurrentSheetName = $sheetName
    $script:CurrentProductName = $selectedProduct

    Set-RunProgress -Stage '检测 PSD 模板' -Detail '正在检查所选 PSD 是否已按智能化套版规范改造。'
    $photoshop = Start-Photoshop
    $preparedPsdPath = Resolve-TemplateForTask -Application $photoshop -TemplatePath $psdPath
    if ($preparedPsdPath -ne $psdPath) {
        $psdPath = $preparedPsdPath
        $script:CurrentPsdPath = $psdPath
        Add-Log "本次任务将使用自动生成的模板副本：$psdPath"
    }

    Set-RunProgress -Stage '建立任务文件夹' -Detail '正在创建本次任务的 JPG、PSD 和任务记录文件夹。'
    $taskOutputDir = New-TaskOutputDirectory -OutputRoot $outputRoot -SheetName $sheetName
    $jpgOutputDir = Join-Path $taskOutputDir 'JPG成品'
    $psdOutputDir = Join-Path $taskOutputDir 'PSD源文件'
    $taskRecordDir = Join-Path $taskOutputDir '任务记录'
    New-Item -Path $jpgOutputDir -ItemType Directory -Force | Out-Null
    New-Item -Path $psdOutputDir -ItemType Directory -Force | Out-Null
    New-Item -Path $taskRecordDir -ItemType Directory -Force | Out-Null
    Assert-WritableDirectory -Path $taskOutputDir -Purpose '任务输出目录'
    Assert-WritableDirectory -Path $jpgOutputDir -Purpose 'JPG 成品输出目录'
    Assert-WritableDirectory -Path $psdOutputDir -Purpose 'PSD 源文件输出目录'
    Assert-WritableDirectory -Path $taskRecordDir -Purpose '任务记录目录'
    $reportPath = Join-Path $taskRecordDir '任务日志.txt'

    Add-Log "Excel：$excelPath"
    Add-Log "PSD：$psdPath"
    Add-Log '素材来源：Excel 图片字段中的完整文件路径（逐条预检，不扫描共享盘）。'
    Add-Log "任务输出目录：$taskOutputDir"

    if (-not (Test-Path -LiteralPath $cleanScript)) { throw "缺少清洗脚本：$cleanScript" }
    if (-not (Test-Path -LiteralPath $sheetScript)) { throw "缺少 Sheet 读取脚本：$sheetScript" }
    if (-not (Test-Path -LiteralPath $batchScript)) { throw "缺少 Photoshop JSX：$batchScript" }

    if ($NoUi) {
        Set-RunProgress -Stage '读取 Sheet 列表' -Detail '正在读取 Excel 中可处理的可见 Sheet。'
        $sheetLines = @(Invoke-Python -Python $python -Arguments @($sheetScript, $excelPath) 2>&1)
        if ($LASTEXITCODE -ne 0) {
            throw "读取 Sheet 失败：$($sheetLines -join '；')"
        }
        $sheets = @($sheetLines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        if ($sheets.Count -eq 0) {
            throw 'Excel 中没有可处理的可见 Sheet。'
        }
        if ([string]::IsNullOrWhiteSpace($SheetName)) {
            throw '命令行模式缺少 -SheetName。'
        }
        if ($sheets -notcontains $SheetName) { throw "命令行指定的 Sheet 不可处理或不存在：$SheetName" }
        $sheetName = $SheetName
        Add-Log "命令行指定 Sheet：$sheetName"
    }
    Add-Log "选择 Sheet：$sheetName"

    if ($NoUi) {
        Set-RunProgress -Stage '读取商品任务列表' -Detail '正在读取当前 Sheet 中的商品文件名。'
        $productLines = @(Invoke-Python -Python $python -Arguments @($sheetScript, '--products', $excelPath, $sheetName) 2>&1)
        if ($LASTEXITCODE -ne 0) {
            throw "读取商品任务列表失败：$($productLines -join '；')"
        }
        $products = @($productLines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        if ([string]::IsNullOrWhiteSpace($ProductName)) {
            throw '命令行模式缺少 -ProductName（全部商品请传 *）。'
        }
        if ($ProductName -eq '*' -or $ProductName -eq '全部商品') {
            $selectedProduct = ''
        } elseif ($products -notcontains $ProductName) {
            throw "命令行指定的商品任务不存在：$ProductName"
        } else {
            $selectedProduct = $ProductName
        }
        Add-Log "命令行指定商品任务：$(if ($selectedProduct) { $selectedProduct } else { '全部商品' })"
    }
    if ([string]::IsNullOrWhiteSpace($selectedProduct)) {
        Add-Log '本次任务：全部商品'
    } else {
        Add-Log "本次任务：$selectedProduct"
    }
    $script:CurrentProductName = $selectedProduct
    Save-UserSettings -ExcelPath $excelPath -PsdPath $psdPath -OutputRoot $outputRoot -SheetName $sheetName -ProductName $selectedProduct
    $taskInfo = @(
        '电商主图套版任务信息',
        '',
        "创建时间：$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))",
        "商品表格：$excelPath",
        "PSD 模板：$psdPath",
        "数据工作表：$sheetName",
        "商品范围：$(if ($selectedProduct) { $selectedProduct } else { '全部商品' })",
        "JPG 成品：$jpgOutputDir",
        "PSD 源文件：$psdOutputDir",
        '',
        '异常商品会自动跳过，具体原因请查看同目录的异常记录.csv。'
    ) -join [Environment]::NewLine
    Write-Utf8Bom -Path (Join-Path $taskRecordDir '任务信息.txt') -Content ($taskInfo + [Environment]::NewLine)

    # All interactive pickers have closed. Showing progress now cannot cover them.
    New-RunProgressWindow
    Set-RunProgress -Stage '数据预检' -Detail '正在校验字段、价格和 Excel 中每个商品素材的完整文件路径。不会扫描共享盘。'
    $cleanArguments = @(
        $cleanScript,
        '--xlsx', $excelPath,
        '--sheet', $sheetName,
        '--output-dir', $taskRecordDir
    )
    if (-not [string]::IsNullOrWhiteSpace($selectedProduct)) {
        foreach ($product in @($selectedProduct -split '\|' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
            $cleanArguments += @('--product', $product)
        }
    }
    if ($Limit -lt 0) {
        throw '命令行参数 -Limit 不能小于 0。'
    }
    if ($Limit -gt 0) {
        if (-not [string]::IsNullOrWhiteSpace($selectedProduct)) {
            throw '命令行参数 -Limit 不能与指定商品任务同时使用。'
        }
        $cleanArguments += @('--limit', $Limit)
        Add-Log "命令行限制本次处理前 $Limit 条有效记录。"
    }
    $cleanOutput = @(Invoke-Python -Python $python -Arguments $cleanArguments 2>&1)
    foreach ($line in $cleanOutput) { Add-Log $line }
    if ($LASTEXITCODE -ne 0) {
        throw "数据清洗失败：$($cleanOutput -join '；')"
    }

    $dataCsv = Join-Path $taskRecordDir 'data.csv'
    $allDataCsv = Join-Path $taskRecordDir 'data_全部记录.csv'
    $errorCsv = Join-Path $taskRecordDir '异常记录.csv'
    if (-not (Test-Path -LiteralPath $dataCsv)) { throw "未生成 data.csv：$dataCsv" }
    if (-not (Test-Path -LiteralPath $allDataCsv)) { throw "未生成 data_全部记录.csv：$allDataCsv" }

    $dataRows = @(Import-Csv -LiteralPath $dataCsv)
    $allDataRows = @(Import-Csv -LiteralPath $allDataCsv)
    $errorRows = if (Test-Path -LiteralPath $errorCsv) { @(Import-Csv -LiteralPath $errorCsv) } else { @() }
    $preflightFailureDetails = (($errorRows | Select-Object -First 10 | ForEach-Object {
        $product = [string]$_.商品文件名
        $detail = if ([string]::IsNullOrWhiteSpace([string]$_.异常详情)) { [string]$_.异常类型 } else { [string]$_.异常详情 }
        "$product：$detail"
    }) -join '；')
    if ($dataRows.Count -eq 0) {
        $issueSummary = if ($errorRows.Count -gt 0) {
            (($errorRows | Group-Object -Property '异常类型' | Sort-Object Count -Descending | ForEach-Object {
                "{0} {1} 条" -f $_.Name, $_.Count
            }) -join '；')
        } else {
            '未生成异常明细'
        }
        $validationMessage = "开始前数据校验未通过：有效记录 0 条，异常 $($errorRows.Count) 条。异常类型：$issueSummary。异常明细：$preflightFailureDetails。请修改 Excel 字段或素材路径后重新运行。异常清单：$errorCsv"
        Add-Log $validationMessage
        Write-Utf8Bom -Path $reportPath -Content (($script:LogLines -join [Environment]::NewLine) + [Environment]::NewLine)
        Close-RunProgressWindow
        Write-EntryStatus -Status 'data_validation_failed' -Message $validationMessage
        Write-TaskHistory -Status '数据未通过' -Message "没有可生成商品；请查看任务记录中的异常记录.csv。"
        exit 1
    }
    $issueSummary = if ($errorRows.Count -gt 0) {
        (($errorRows | Group-Object -Property '异常类型' | Sort-Object Count -Descending | ForEach-Object {
            "{0} {1} 条" -f $_.Name, $_.Count
        }) -join '；')
    } else {
        '无'
    }
    Add-Log "数据校验完成：可处理记录 $($dataRows.Count) 条；异常 $($errorRows.Count) 条；异常类型：$issueSummary。异常清单：$errorCsv"
    $preflightMode = 'valid_only'
    if ($errorRows.Count -gt 0 -and -not $NoUi) {
        Add-Log "已自动跳过 $($errorRows.Count) 条异常商品；详情仅写入任务记录，不弹出逐条错误。"
        Set-RunProgress -Stage '启动 Photoshop' -Detail "正在处理 $($dataRows.Count) 条无问题商品；异常商品已记录并跳过。"
    }
    $priceSingleRows = @($dataRows | Where-Object {
        $_.PSObject.Properties.Name -contains '价格' -and
        -not [string]::IsNullOrWhiteSpace($_.价格) -and
        -not [string]::IsNullOrWhiteSpace($_.价格1) -and
        -not [string]::IsNullOrWhiteSpace($_.价格2)
    })
    if ($priceSingleRows.Count -gt 0) {
        Add-Log "任务级提示 W_PRICE_UNCONFIRMED：发现 $($priceSingleRows.Count) 条记录由【价格】单值拆分为【价格1/价格2】，请业务复核价格口径；该提示仅汇总一次，不阻断套版。"
    }

    Set-RunProgress -Stage '启动 Photoshop' -Detail '数据预检通过，已连接 Photoshop，准备打开模板。'
    Set-RunProgress -Stage '打开 PSD 模板' -Detail '正在打开模板。Photoshop 出现后请勿关闭。'
    $script:OpenedDocument = $photoshop.Open($psdPath)
    Start-Sleep -Seconds 2

    Set-RunProgress -Stage 'Photoshop 正在导出' -Detail "正在处理 $($dataRows.Count) 条商品。Photoshop 内会显示当前进度。"
    # Only measure the actual Photoshop replacement/export work. Template
    # checks and Excel preflight are intentionally excluded from this time.
    $generationStartedAt = Get-Date
    $batchText = [System.IO.File]::ReadAllText($batchScript, [System.Text.Encoding]::UTF8)
    $prefix = @(
        '$.global.__BATCH_INPUTS__ = {',
        ('  csv: {0},' -f (ConvertTo-JsStringLiteral $dataCsv)),
        ('  output: {0},' -f (ConvertTo-JsStringLiteral $jpgOutputDir)),
        ('  psdOutput: {0},' -f (ConvertTo-JsStringLiteral $psdOutputDir)),
        ('  continueWithPreflightIssues: {0}' -f ($(if ($preflightMode -eq 'all_rows') { 'true' } else { 'false' }))),
        '};'
    ) -join "`r`n"
    $jsxText = $prefix + "`r`n" + $batchText
    $jsxResult = Invoke-PhotoshopJavaScript -Application $photoshop -ScriptText $jsxText
    # Photoshop may return its restored DialogModes enum here. It is not an
    # export error and should not be shown as one in the designer-facing log.
    if ($jsxResult -and [string]$jsxResult -ne 'DialogModes.ERROR') {
        Add-Log "Photoshop 返回：$jsxResult"
    }

    $resultReport = Join-Path $jpgOutputDir '结果报告.csv'
    $resultSummary = Get-PhotoshopResultSummary -ResultReport $resultReport -JpgOutputDir $jpgOutputDir
    $storedResultReport = Join-Path $taskRecordDir '生成结果.csv'
    Move-Item -LiteralPath $resultReport -Destination $storedResultReport -Force
    $resultReport = $storedResultReport
    Close-OpenedPhotoshopDocument

    if ($errorRows.Count -gt 0 -and $preflightMode -ne 'all_rows') {
        if ($resultSummary.Outcome -ne 'failed') {
            $resultSummary.Outcome = 'needs_review'
        }
        $resultSummary.ReviewCount += $errorRows.Count
        $resultSummary.FailureCount += $errorRows.Count
        $resultSummary.FailureDetails = (@($preflightFailureDetails, $resultSummary.FailureDetails | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join '；')
        $resultSummary.SummaryText += " 数据预检拦截：$($errorRows.Count) 条。"
    }

    Add-Log "Photoshop 结果汇总：$($resultSummary.SummaryText)结果报告：$resultReport"
    if ($resultSummary.Outcome -eq 'failed') {
        $criticalText = if ([string]::IsNullOrWhiteSpace($resultSummary.CriticalCodes)) { '无' } else { $resultSummary.CriticalCodes }
        throw "Photoshop 结果失败：$($resultSummary.SummaryText)关键错误码：$criticalText。结果报告：$resultReport"
    }

    if ($resultSummary.Outcome -eq 'needs_review') {
        Add-Log '任务状态：需复核。'
        $elapsedText = Get-TaskElapsedText -StartedAt $generationStartedAt
        Add-Log $elapsedText
        Write-Utf8Bom -Path $reportPath -Content (($script:LogLines -join [Environment]::NewLine) + [Environment]::NewLine)
        Close-RunProgressWindow
        $reviewMessage = "套版处理完成，存在需复核条目。`r`n$($resultSummary.SummaryText)`r`n$elapsedText"
        if (-not [string]::IsNullOrWhiteSpace($resultSummary.FailureDetails)) {
            $reviewMessage += "`r`n失败明细：$($resultSummary.FailureDetails)"
        }
        Write-EntryStatus -Status 'needs_review' -Message $reviewMessage
        Write-TaskHistory -Status '已完成，有跳过项' -Message "已导出 $($resultSummary.ExportedCount) 张；跳过或需核对 $($resultSummary.FailureCount) 条。"
        if (-not $NoUi) {
            Show-TaskCompletionDialog -Title '套版已完成' -Message "套版完成。`r`n已生成 $($resultSummary.ExportedCount) 张 JPG。`r`n有 $($resultSummary.FailureCount) 条未生成，详情已保存在任务记录中。`r`n$elapsedText" -JpgOutputDir $jpgOutputDir
        }
        exit 0
    }

    Add-Log '任务状态：完成。'
    $elapsedText = Get-TaskElapsedText -StartedAt $generationStartedAt
    Add-Log $elapsedText
    Write-Utf8Bom -Path $reportPath -Content (($script:LogLines -join [Environment]::NewLine) + [Environment]::NewLine)
    Close-RunProgressWindow
    Write-EntryStatus -Status 'success' -Message "套版完成。$($resultSummary.SummaryText) $elapsedText"
    Write-TaskHistory -Status '已完成' -Message "已导出 $($resultSummary.ExportedCount) 张 JPG。$elapsedText"
    if (-not $NoUi) {
            Show-TaskCompletionDialog -Title '套版已完成' -Message "套版完成。`r`n$($resultSummary.SummaryText)`r`n$elapsedText" -JpgOutputDir $jpgOutputDir
    }
    exit 0
} catch {
    $message = "任务失败：$($_.Exception.Message)"
    Add-Log $message
    Close-RunProgressWindow
    Close-OpenedPhotoshopDocument
    if ($reportPath) {
        Write-Utf8Bom -Path $reportPath -Content (($script:LogLines -join [Environment]::NewLine) + [Environment]::NewLine)
    }
    Write-EntryFailureReport -ErrorSummary $message -Suggestion '错误已经写入工具任务记录；需要排查时，请提供任务文件夹中的任务记录，或用户目录下的工具任务记录。'
    Write-TaskHistory -Status '失败' -Message $message
    exit 1
}
