$ErrorActionPreference = 'Stop'

function New-Step {
    param(
        [string]$Name,
        [string]$Status = '未执行',
        [string]$Detail = ''
    )

    [pscustomobject]@{
        Name = $Name
        Status = $Status
        Detail = $Detail
    }
}

function Set-Step {
    param(
        [string]$Name,
        [string]$Status,
        [string]$Detail = ''
    )

    $script:Steps[$Name].Status = $Status
    $script:Steps[$Name].Detail = $Detail
}

function Add-ErrorSummary {
    param([string]$Message)

    if (-not [string]::IsNullOrWhiteSpace($Message)) {
        $script:Errors.Add($Message) | Out-Null
    }
}

function ConvertTo-CsvValue {
    param([string]$Value)

    if ($null -eq $Value) {
        $Value = ''
    }

    '"' + ($Value -replace '"', '""') + '"'
}

function Save-Utf8Bom {
    param(
        [string]$Path,
        [string]$Content
    )

    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($Path, $Content, $utf8Bom)
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

function Get-PhotoshopRegistryPaths {
    $paths = New-Object System.Collections.Generic.List[string]

    $appPathRoots = @(
        'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Photoshop.exe',
        'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\Photoshop.exe',
        'Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Photoshop.exe'
    )

    foreach ($root in $appPathRoots) {
        $value = Get-RegistryValueSafe -Path $root -Name '(default)'
        if ($value -and (Test-Path -LiteralPath $value)) {
            $paths.Add($value) | Out-Null
        }
    }

    $adobeRoots = @(
        'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Adobe\Photoshop',
        'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Adobe\Photoshop',
        'Registry::HKEY_CURRENT_USER\SOFTWARE\Adobe\Photoshop'
    )

    foreach ($root in $adobeRoots) {
        try {
            $versions = Get-ChildItem -Path $root -ErrorAction Stop
            foreach ($version in $versions) {
                $appDir = Get-RegistryValueSafe -Path $version.PSPath -Name 'ApplicationPath'
                if ($appDir) {
                    $candidate = Join-Path -Path $appDir -ChildPath 'Photoshop.exe'
                    if (Test-Path -LiteralPath $candidate) {
                        $paths.Add($candidate) | Out-Null
                    }
                }
            }
        } catch {
        }
    }

    return $paths | Select-Object -Unique
}

function Get-PhotoshopComProgIds {
    $progIds = New-Object System.Collections.Generic.List[string]
    $progIds.Add('Photoshop.Application') | Out-Null

    $classRoots = @(
        'Registry::HKEY_CLASSES_ROOT',
        'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Classes',
        'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Classes',
        'Registry::HKEY_CURRENT_USER\SOFTWARE\Classes'
    )

    foreach ($root in $classRoots) {
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

function Invoke-PhotoshopJavaScript {
    param(
        [object]$Application,
        [string]$ScriptText,
        [string]$MarkerPath
    )

    try {
        return $Application.DoJavaScript($ScriptText, @($MarkerPath), 2)
    } catch {
        $firstError = $_.Exception.Message
    }

    try {
        return $Application.DoJavaScript($ScriptText, @($MarkerPath))
    } catch {
        throw "DoJavaScript 调用失败：$firstError / $($_.Exception.Message)"
    }
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$outputDir = Join-Path -Path $scriptDir -ChildPath '输出'
$jsxPath = Join-Path -Path $scriptDir -ChildPath 'M1_最小验证.jsx'
$txtReportPath = Join-Path -Path $outputDir -ChildPath 'M1验证报告.txt'
$csvReportPath = Join-Path -Path $outputDir -ChildPath 'M1验证报告.csv'
$markerPath = Join-Path -Path $outputDir -ChildPath ("M1_Photoshop_COM_marker_{0}.txt" -f ([Guid]::NewGuid().ToString('N')))

$script:Errors = New-Object System.Collections.Generic.List[string]
$script:Steps = @{}
foreach ($stepName in @('Windows检测', 'PowerShell检测', 'Photoshop定位', 'COM启动', 'JSX执行', '标记回传')) {
    $script:Steps[$stepName] = New-Step -Name $stepName
}

$result = [ordered]@{
    Status = '未通过'
    Time = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    WindowsVersion = ''
    PowerShellVersion = $PSVersionTable.PSVersion.ToString()
    PhotoshopPath = ''
    PhotoshopVersion = ''
    ComProgId = ''
    ComStarted = '否'
    JsxExecuted = '否'
    MarkerReturned = '否'
    MarkerPath = $markerPath
    MarkerPathMode = ''
    ErrorSummary = ''
    Suggestion = ''
}

try {
    if (-not (Test-Path -LiteralPath $outputDir)) {
        New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
    }

    Write-Host '正在检测 Windows 环境...'
    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
    $result.WindowsVersion = "$($os.Caption) $($os.Version) Build $($os.BuildNumber)"
    Set-Step -Name 'Windows检测' -Status '通过' -Detail $result.WindowsVersion
} catch {
    $result.WindowsVersion = [System.Environment]::OSVersion.VersionString
    Set-Step -Name 'Windows检测' -Status '未通过' -Detail $_.Exception.Message
    Add-ErrorSummary "Windows 检测失败：$($_.Exception.Message)"
}

try {
    Write-Host '正在检测 PowerShell 版本...'
    if (($PSVersionTable.PSVersion.Major -lt 5) -or (($PSVersionTable.PSVersion.Major -eq 5) -and ($PSVersionTable.PSVersion.Minor -lt 1))) {
        throw "当前 PowerShell 版本为 $($PSVersionTable.PSVersion)，需要 Windows PowerShell 5.1 或更高版本。"
    }
    Set-Step -Name 'PowerShell检测' -Status '通过' -Detail $result.PowerShellVersion
} catch {
    Set-Step -Name 'PowerShell检测' -Status '未通过' -Detail $_.Exception.Message
    Add-ErrorSummary "PowerShell 检测失败：$($_.Exception.Message)"
}

try {
    Write-Host '正在定位 Photoshop...'
    $registryPaths = @(Get-PhotoshopRegistryPaths)
    if ($registryPaths.Count -gt 0) {
        $result.PhotoshopPath = $registryPaths[0]
        $pathStatus = '通过'
    } else {
        $result.PhotoshopPath = '未在注册表中找到 Photoshop.exe；继续尝试 COM ProgID'
        $pathStatus = '部分通过'
    }

    $progIds = @(Get-PhotoshopComProgIds)
    Set-Step -Name 'Photoshop定位' -Status $pathStatus -Detail ("COM ProgID: {0}; 路径: {1}" -f ($progIds -join ', '), $result.PhotoshopPath)
} catch {
    Set-Step -Name 'Photoshop定位' -Status '未通过' -Detail $_.Exception.Message
    Add-ErrorSummary "Photoshop 定位失败：$($_.Exception.Message)"
    $progIds = @('Photoshop.Application')
}

$photoshop = $null
try {
    Write-Host '正在通过 COM 启动 Photoshop...'
    foreach ($progId in $progIds) {
        try {
            $photoshop = New-Object -ComObject $progId -ErrorAction Stop
            $result.ComProgId = $progId
            break
        } catch {
            Add-ErrorSummary "COM ProgID $progId 启动失败：$($_.Exception.Message)"
        }
    }

    if ($null -eq $photoshop) {
        throw '所有 Photoshop COM ProgID 均启动失败。'
    }

    $photoshop.Visible = $true
    Start-Sleep -Seconds 3
    $result.ComStarted = '是'
    Set-Step -Name 'COM启动' -Status '通过' -Detail ("ProgID: {0}" -f $result.ComProgId)

    try {
        $result.PhotoshopVersion = [string]$photoshop.Version
    } catch {
        $result.PhotoshopVersion = '已启动，但无法读取版本'
    }
} catch {
    Set-Step -Name 'COM启动' -Status '未通过' -Detail $_.Exception.Message
    Add-ErrorSummary "COM 启动失败：$($_.Exception.Message)"
}

try {
    if ($null -eq $photoshop) {
        throw 'Photoshop COM 未启动，跳过 JSX 执行。'
    }

    if (-not (Test-Path -LiteralPath $jsxPath)) {
        throw "未找到 JSX 文件：$jsxPath"
    }

    Write-Host '正在执行最小 JSX 验证...'

    $jsxText = [System.IO.File]::ReadAllText($jsxPath, [System.Text.Encoding]::UTF8)
    $null = Invoke-PhotoshopJavaScript -Application $photoshop -ScriptText $jsxText -MarkerPath $markerPath
    Start-Sleep -Seconds 2

    $result.JsxExecuted = '是'
    Set-Step -Name 'JSX执行' -Status '通过' -Detail 'DoJavaScript 已返回'
} catch {
    Set-Step -Name 'JSX执行' -Status '未通过' -Detail $_.Exception.Message
    Add-ErrorSummary "JSX 执行失败：$($_.Exception.Message)"
}

try {
    Write-Host '正在检查 JSX 回传标记...'
    if (Test-Path -LiteralPath $markerPath) {
        $result.MarkerPathMode = '输出目录GUID marker'
    } else {
        throw '未找到 JSX 写出的标记文件。'
    }

    $markerText = [System.IO.File]::ReadAllText($markerPath, [System.Text.Encoding]::UTF8)
    if ($markerText -notmatch '^M1_RESULT\|成功\|') {
        throw "标记文件内容不符合预期：$markerText"
    }

    $result.MarkerReturned = '是'
    Set-Step -Name '标记回传' -Status '通过' -Detail ("{0}；路径：{1}；模式：{2}" -f $markerText.Trim(), $result.MarkerPath, $result.MarkerPathMode)
} catch {
    Set-Step -Name '标记回传' -Status '未通过' -Detail $_.Exception.Message
    Add-ErrorSummary "标记回传失败：$($_.Exception.Message)"
}

if (($script:Steps['PowerShell检测'].Status -eq '通过') -and ($result.ComStarted -eq '是') -and ($result.JsxExecuted -eq '是') -and ($result.MarkerReturned -eq '是')) {
    $result.Status = '通过'
    $result.Suggestion = 'M1 验证通过：本机可以通过 Windows COM 调用 Photoshop 并执行 JSX。'
} else {
    $result.Status = '未通过'
    if ($script:Steps['PowerShell检测'].Status -ne '通过') {
        $result.Suggestion = '请使用 Windows PowerShell 5.1 或更高版本运行验证包，并把报告交给开发同事确认入口环境。'
    } elseif ($result.ComStarted -ne '是') {
        $result.Suggestion = '请确认已安装 Windows 版 Photoshop，并能正常手动打开；如仍失败，请把报告交给开发检查 COM 注册。'
    } elseif ($result.JsxExecuted -ne '是') {
        $result.Suggestion = '请把报告交给开发检查 Photoshop DoJavaScript/JSX 执行权限或版本兼容性。'
    } else {
        $result.Suggestion = '请把报告交给开发检查 JSX 参数传递、输出目录写入权限或中文路径兼容性。'
    }
}

if ($script:Errors.Count -gt 0) {
    $result.ErrorSummary = ($script:Errors | Select-Object -Unique) -join '；'
} else {
    $result.ErrorSummary = '无'
}

$lines = New-Object System.Collections.Generic.List[string]
$lines.Add('M1 Windows Photoshop COM 验证报告') | Out-Null
$lines.Add('') | Out-Null
$lines.Add("状态：$($result.Status)") | Out-Null
$lines.Add("时间：$($result.Time)") | Out-Null
$lines.Add("Windows版本：$($result.WindowsVersion)") | Out-Null
$lines.Add("PowerShell版本：$($result.PowerShellVersion)") | Out-Null
$lines.Add("Photoshop路径：$($result.PhotoshopPath)") | Out-Null
$lines.Add("Photoshop版本：$($result.PhotoshopVersion)") | Out-Null
$lines.Add("COM ProgID：$($result.ComProgId)") | Out-Null
$lines.Add("COM启动：$($result.ComStarted)") | Out-Null
$lines.Add("JSX执行：$($result.JsxExecuted)") | Out-Null
$lines.Add("标记回传：$($result.MarkerReturned)") | Out-Null
$lines.Add("标记文件路径：$($result.MarkerPath)") | Out-Null
$lines.Add("标记路径模式：$($result.MarkerPathMode)") | Out-Null
$lines.Add("错误摘要：$($result.ErrorSummary)") | Out-Null
$lines.Add("建议动作：$($result.Suggestion)") | Out-Null
$lines.Add('') | Out-Null
$lines.Add('分项结果：') | Out-Null
foreach ($step in $script:Steps.Values) {
    $lines.Add("- $($step.Name)：$($step.Status)；$($step.Detail)") | Out-Null
}

Save-Utf8Bom -Path $txtReportPath -Content (($lines -join [Environment]::NewLine) + [Environment]::NewLine)

$csvRows = New-Object System.Collections.Generic.List[string]
$csvRows.Add('字段,值') | Out-Null
foreach ($key in $result.Keys) {
    $csvRows.Add(('{0},{1}' -f (ConvertTo-CsvValue $key), (ConvertTo-CsvValue ([string]$result[$key])))) | Out-Null
}
foreach ($step in $script:Steps.Values) {
    $csvRows.Add(('{0},{1}' -f (ConvertTo-CsvValue ("步骤-$($step.Name)")), (ConvertTo-CsvValue ("$($step.Status)；$($step.Detail)")))) | Out-Null
}

Save-Utf8Bom -Path $csvReportPath -Content (($csvRows -join [Environment]::NewLine) + [Environment]::NewLine)

Write-Host ''
Write-Host "状态：$($result.Status)"
Write-Host "报告已生成：$txtReportPath"
Write-Host "CSV已生成：$csvReportPath"

try {
    Start-Process -FilePath $outputDir
} catch {
    Write-Host "无法自动打开输出文件夹，请手动打开：$outputDir"
}

if ($result.Status -eq '通过') {
    exit 0
}

exit 1
