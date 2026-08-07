param(
    [Parameter(Mandatory = $true)][string]$ToolRoot,
    [Parameter(Mandatory = $true)][string]$ExcelPath,
    [Parameter(Mandatory = $true)][string]$PsdPath,
    [Parameter(Mandatory = $true)][string]$OutputRoot,
    [Parameter(Mandatory = $true)][string]$ArtifactDir,
    [string]$Category = '',
    [string]$Channel = '',
    [string]$SheetName = '',
    [int]$ExpectedJpgWidth = 800,
    [int]$ExpectedJpgHeight = 800,
    [int]$ProductCount = 0,
    [switch]$UseSingleProduct,
    [switch]$RequireTextOverflow
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class NativeMouse {
    [StructLayout(LayoutKind.Sequential)]
    private struct Point {
        public int X;
        public int Y;
    }

    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint flags, uint dx, uint dy, uint data, UIntPtr extraInfo);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr window);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr window, int command);
    [DllImport("user32.dll")] private static extern bool ScreenToClient(IntPtr window, ref Point point);
    [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);
    public const uint LeftDown = 0x0002;
    public const uint LeftUp = 0x0004;

    public static bool ClickClientPoint(IntPtr window, int screenX, int screenY) {
        Point point = new Point { X = screenX, Y = screenY };
        if (!ScreenToClient(window, ref point)) return false;

        const uint WmLeftButtonDown = 0x0201;
        const uint WmLeftButtonUp = 0x0202;
        IntPtr position = new IntPtr(unchecked((point.Y << 16) | (point.X & 0xffff)));
        SendMessage(window, WmLeftButtonDown, new IntPtr(1), position);
        SendMessage(window, WmLeftButtonUp, IntPtr.Zero, position);
        return true;
    }
}
'@

New-Item -Path $ArtifactDir -ItemType Directory -Force | Out-Null
$logPath = Join-Path $ArtifactDir 'ui-smoke.log'
$process = $null
$existingTaskDirectoryPaths = @()

function Write-SmokeLog {
    param([string]$Message)
    Add-Content -LiteralPath $logPath -Value ("[{0}] {1}" -f (Get-Date).ToString('HH:mm:ss'), $Message) -Encoding UTF8
}

function Get-DesktopWindow {
    param([string]$Title)
    $condition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::NameProperty,
        $Title
    )
    return [System.Windows.Automation.AutomationElement]::RootElement.FindFirst(
        [System.Windows.Automation.TreeScope]::Children,
        $condition
    )
}

function Wait-DesktopWindow {
    param(
        [string]$Title,
        [int]$TimeoutSeconds = 30
    )
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        $window = Get-DesktopWindow -Title $Title
        if ($window) { return $window }
        Start-Sleep -Milliseconds 500
    } while ((Get-Date) -lt $deadline)
    throw "等待窗口超时：$Title"
}

function Get-Control {
    param(
        [System.Windows.Automation.AutomationElement]$Root,
        [string]$Name,
        [System.Windows.Automation.ControlType]$ControlType
    )
    $conditions = @(
        (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, $Name)),
        (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty, $ControlType))
    )
    $condition = New-Object System.Windows.Automation.AndCondition($conditions)
    return $Root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $condition)
}

function Invoke-Control {
    param([System.Windows.Automation.AutomationElement]$Control)
    if (-not $Control) { throw '未找到要点击的控件。' }
    $pattern = $Control.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
    $pattern.Invoke()
}

function Select-ComboItem {
    param([System.Windows.Automation.AutomationElement]$Combo, [string]$Name)
    if (-not $Combo) { throw "未找到下拉框：$Name" }
    $Combo.GetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern).Expand()
    $condition = New-Object System.Windows.Automation.AndCondition(@(
        (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, $Name)),
        (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::ListItem))
    ))
    $deadline = (Get-Date).AddSeconds(5)
    do {
        $item = [System.Windows.Automation.AutomationElement]::RootElement.FindFirst(
            [System.Windows.Automation.TreeScope]::Subtree,
            $condition
        )
        if ($item) {
            $item.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select()
            Start-Sleep -Milliseconds 250
            return
        }
        Start-Sleep -Milliseconds 250
    } while ((Get-Date) -lt $deadline)
    throw "下拉框没有选项：$Name"
}

function Show-AutomationWindow {
    param([System.Windows.Automation.AutomationElement]$Window)
    if (-not $Window) { return }
    $handle = [IntPtr]$Window.Current.NativeWindowHandle
    try {
        $windowPattern = $Window.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern)
        if ($windowPattern.Current.WindowVisualState -eq [System.Windows.Automation.WindowVisualState]::Minimized) {
            $windowPattern.SetWindowVisualState([System.Windows.Automation.WindowVisualState]::Normal)
        }
    } catch {
        [void][NativeMouse]::ShowWindow($handle, 9)
    }
    [void][NativeMouse]::SetForegroundWindow($handle)
    Start-Sleep -Milliseconds 250
}

function Click-Control {
    param(
        [System.Windows.Automation.AutomationElement]$Control,
        [switch]$ClickLeft
    )
    if (-not $Control) { throw '未找到要点击的控件。' }
    $window = $Control
    $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
    while ($window -and $window.Current.ControlType -ne [System.Windows.Automation.ControlType]::Window) {
        $window = $walker.GetParent($window)
    }
    if ($window) {
        Show-AutomationWindow -Window $window
    }
    $rectangle = $Control.Current.BoundingRectangle
    $x = if ($ClickLeft) { [int]($rectangle.X + 9) } else { [int]($rectangle.X + ($rectangle.Width / 2)) }
    $y = [int]($rectangle.Y + ($rectangle.Height / 2))
    [void][NativeMouse]::SetCursorPos($x, $y)
    [NativeMouse]::mouse_event([NativeMouse]::LeftDown, 0, 0, 0, [UIntPtr]::Zero)
    [NativeMouse]::mouse_event([NativeMouse]::LeftUp, 0, 0, 0, [UIntPtr]::Zero)
}

function Set-ControlValue {
    param(
        [System.Windows.Automation.AutomationElement]$Control,
        [string]$Value
    )
    $pattern = $Control.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern)
    $pattern.SetValue($Value)
}

function Set-CheckedListItem {
    param([System.Windows.Automation.AutomationElement]$Item)
    if (-not $Item) { throw '未找到要勾选的商品。' }

    # CheckedListBox keeps selection and check state separately. Send a click
    # directly to the list HWND so the action does not depend on focus or the
    # keyboard state. The production submit action below verifies the check
    # state through the same path used by an operator.
    $rectangle = $Item.Current.BoundingRectangle
    if ($rectangle.Width -le 0 -or $rectangle.Height -le 0) {
        throw '商品复选框不可见，无法勾选。'
    }

    $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
    $list = $walker.GetParent($Item)
    if (-not $list -or $list.Current.ControlType -ne [System.Windows.Automation.ControlType]::List) {
        throw '无法定位商品复选列表控件。'
    }
    $listHandle = [IntPtr]$list.Current.NativeWindowHandle
    if ($listHandle -eq [IntPtr]::Zero) {
        throw '商品复选列表没有可用的原生窗口句柄。'
    }
    if (-not [NativeMouse]::ClickClientPoint(
        $listHandle,
        [int]($rectangle.X + 8),
        [int]($rectangle.Y + ($rectangle.Height / 2))
    )) {
        throw '商品复选框坐标转换失败。'
    }

    Write-SmokeLog '已向商品复选框发送稳定原生点击，等待生产入口验证勾选状态。'
}

function Capture-Desktop {
    param([string]$Name)
    $bounds = [System.Windows.Forms.SystemInformation]::VirtualScreen
    $bitmap = New-Object System.Drawing.Bitmap($bounds.Width, $bounds.Height)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    try {
        $graphics.CopyFromScreen($bounds.Left, $bounds.Top, 0, 0, $bitmap.Size)
        $bitmap.Save((Join-Path $ArtifactDir $Name), [System.Drawing.Imaging.ImageFormat]::Png)
    } finally {
        $graphics.Dispose()
        $bitmap.Dispose()
    }
}

function Format-BoundsValue {
    param([double]$Value)
    if ([double]::IsNaN($Value) -or [double]::IsInfinity($Value)) {
        return 'offscreen'
    }
    if ($Value -gt [int]::MaxValue -or $Value -lt [int]::MinValue) {
        return 'offscreen'
    }
    return [string][Math]::Round($Value)
}

function Write-ControlSnapshot {
    param(
        [System.Windows.Automation.AutomationElement]$Window,
        [string]$Label
    )
    Write-SmokeLog "控件快照：$Label"
    $controls = $Window.FindAll(
        [System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.Condition]::TrueCondition
    )
    foreach ($control in $controls) {
        $rectangle = $control.Current.BoundingRectangle
        Write-SmokeLog ("{0} | {1} | x={2} y={3} w={4} h={5}" -f $control.Current.ControlType.ProgrammaticName, $control.Current.Name, (Format-BoundsValue $rectangle.X), (Format-BoundsValue $rectangle.Y), (Format-BoundsValue $rectangle.Width), (Format-BoundsValue $rectangle.Height))
    }
}

try {
    Write-SmokeLog '开始 Windows 界面回归。'
    if (Test-Path -LiteralPath $OutputRoot -PathType Container) {
        $existingTaskDirectoryPaths = @(
            Get-ChildItem -LiteralPath $OutputRoot -Directory -ErrorAction SilentlyContinue |
                Select-Object -ExpandProperty FullName
        )
    }
    $entry = Join-Path $ToolRoot '开始套版.cmd'
    if (-not (Test-Path -LiteralPath $entry -PathType Leaf)) {
        # Windows tar can decode a Unicode archive member with the active
        # console code page. The ASCII launcher is equivalent and keeps the
        # smoke path independent of archive filename decoding.
        $entry = Join-Path $ToolRoot 'L0_Start.cmd'
    }
    if (-not (Test-Path -LiteralPath $entry -PathType Leaf)) {
        throw "未找到用户启动入口：$entry"
    }
    $arguments = '/d /c call "{0}"' -f $entry
    $process = Start-Process -FilePath (Join-Path $env:SystemRoot 'System32\cmd.exe') -ArgumentList $arguments -WorkingDirectory $ToolRoot -WindowStyle Normal -PassThru
    Start-Sleep -Milliseconds 750
    Write-SmokeLog ("入口进程：PID=$($process.Id)，已退出=$($process.HasExited)，工作目录=$ToolRoot")
    if ($process.HasExited) {
        throw "工具入口启动后立即退出，退出码：$($process.ExitCode)"
    }

    $channelWindow = Wait-DesktopWindow -Title '选择品类和渠道' -TimeoutSeconds 45
    Show-AutomationWindow -Window $channelWindow
    if ($Category -or $Channel) {
        if (-not $Category -or -not $Channel) { throw '品类和渠道必须同时指定。' }
        $combos = @($channelWindow.FindAll(
            [System.Windows.Automation.TreeScope]::Descendants,
            (New-Object System.Windows.Automation.PropertyCondition(
                [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
                [System.Windows.Automation.ControlType]::ComboBox
            ))
        ) | Sort-Object { $_.Current.BoundingRectangle.Y })
        if ($combos.Count -ne 2) { throw "品类渠道下拉框数量异常：$($combos.Count)" }
        Select-ComboItem -Combo $combos[0] -Name $Category
        Select-ComboItem -Combo $combos[1] -Name $Channel
    }
    Write-ControlSnapshot -Window $channelWindow -Label '品类渠道选择'
    Capture-Desktop -Name '00-品类渠道选择.png'
    Invoke-Control -Control (Get-Control -Root $channelWindow -Name '下一步' -ControlType ([System.Windows.Automation.ControlType]::Button))

    $mainWindow = Wait-DesktopWindow -Title '电商主图套版工具 1.4' -TimeoutSeconds 45
    Show-AutomationWindow -Window $mainWindow
    $photoshopHint = Get-Control -Root $mainWindow -Name '请先启动并登录 Photoshop，进入首页后再选择商品表格和 PSD 模板。' -ControlType ([System.Windows.Automation.ControlType]::Text)
    if (-not $photoshopHint) { throw '初始页缺少 Photoshop 启动提醒。' }
    Write-ControlSnapshot -Window $mainWindow -Label '初始页'
    Capture-Desktop -Name '01-初始页.png'

    $edits = @($mainWindow.FindAll(
        [System.Windows.Automation.TreeScope]::Descendants,
        (New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
            [System.Windows.Automation.ControlType]::Edit
        ))) | Sort-Object { $_.Current.BoundingRectangle.Y })
    if ($edits.Count -lt 3) { throw "输入框数量异常：$($edits.Count)" }
    Set-ControlValue -Control $edits[0] -Value $ExcelPath
    Set-ControlValue -Control $edits[1] -Value $PsdPath
    Set-ControlValue -Control $edits[2] -Value $OutputRoot
    Invoke-Control -Control (Get-Control -Root $mainWindow -Name '重新读取' -ControlType ([System.Windows.Automation.ControlType]::Button))
    Start-Sleep -Seconds 5

    $mainWindow = Wait-DesktopWindow -Title '电商主图套版工具 1.4'
    Show-AutomationWindow -Window $mainWindow
    if (-not [string]::IsNullOrWhiteSpace($SheetName)) {
        $sheetCombo = Get-Control -Root $mainWindow -Name '3  数据工作表' -ControlType ([System.Windows.Automation.ControlType]::ComboBox)
        if (-not $sheetCombo) { throw '未找到数据工作表下拉框。' }
        Select-ComboItem -Combo $sheetCombo -Name $SheetName
        Start-Sleep -Seconds 3
        $mainWindow = Wait-DesktopWindow -Title '电商主图套版工具 1.4'
        Show-AutomationWindow -Window $mainWindow
        Write-SmokeLog "已显式选择 Sheet：$SheetName"
    }
    Write-ControlSnapshot -Window $mainWindow -Label '表格已加载'
    Capture-Desktop -Name '02-表格已加载.png'

    $partialRadio = Get-Control -Root $mainWindow -Name '只生成勾选的商品' -ControlType ([System.Windows.Automation.ControlType]::RadioButton)
    Click-Control -Control $partialRadio
    Start-Sleep -Milliseconds 500
    $productListControl = Get-Control -Root $mainWindow -Name '4  商品范围' -ControlType ([System.Windows.Automation.ControlType]::List)
    if (-not $productListControl) { throw '商品列表为空。' }
    Show-AutomationWindow -Window $mainWindow
    $productItems = @($productListControl.FindAll(
        [System.Windows.Automation.TreeScope]::Children,
        (New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
            [System.Windows.Automation.ControlType]::ListItem
        ))
    ))
    if ($productItems.Count -lt 1) { throw '商品列表没有可选商品。' }
    $selectionCount = if ($ProductCount -gt 0) { $ProductCount } else { 1 }
    if ($selectionCount -gt $productItems.Count) {
        throw "请求勾选 $selectionCount 个商品，但当前列表只有 $($productItems.Count) 个。"
    }
    for ($selectionIndex = 0; $selectionIndex -lt $selectionCount; $selectionIndex++) {
        Set-CheckedListItem -Item $productItems[$selectionIndex]
        Start-Sleep -Milliseconds 250
    }
    Write-SmokeLog "已勾选 $selectionCount 个商品。"
    Start-Sleep -Milliseconds 500
    Capture-Desktop -Name '03-商品已勾选.png'

    if (-not $UseSingleProduct -and $ProductCount -le 0) {
        # The checked-item state is visually verified before switching back to
        # the full-sheet path used by the default regression.
        $allRadio = Get-Control -Root $mainWindow -Name '全部商品' -ControlType ([System.Windows.Automation.ControlType]::RadioButton)
        Click-Control -Control $allRadio
        Start-Sleep -Milliseconds 500
    }

    $startButton = Get-Control -Root $mainWindow -Name '开始生成' -ControlType ([System.Windows.Automation.ControlType]::Button)
    Invoke-Control -Control $startButton
    if ($UseSingleProduct) {
        Start-Sleep -Milliseconds 800
        if (Get-DesktopWindow -Title '电商主图套版工具 1.4') {
            throw '生产入口仍停留在商品范围页，复选框未被实际勾选。'
        }
        Write-SmokeLog '生产入口已接受单商品选择，复选框状态验证通过。'
    }
    Write-SmokeLog '已点击开始生成。'

    $deadline = (Get-Date).AddMinutes(4)
    $completionWindow = $null
    do {
        $templatePrompt = Get-DesktopWindow -Title '主图模板检查'
        if ($templatePrompt) {
            $buttons = @($templatePrompt.FindAll(
                [System.Windows.Automation.TreeScope]::Descendants,
                (New-Object System.Windows.Automation.PropertyCondition(
                    [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
                    [System.Windows.Automation.ControlType]::Button
                ))))
            $yesButton = $buttons | Where-Object { $_.Current.Name -match '^是|Yes' } | Select-Object -First 1
            if ($yesButton) { Invoke-Control -Control $yesButton }
        }
        $completionWindow = Get-DesktopWindow -Title '套版已完成'
        if ($completionWindow) { break }
        if ($process.HasExited) { throw "工具在完成页出现前退出，退出码：$($process.ExitCode)" }
        Start-Sleep -Seconds 1
    } while ((Get-Date) -lt $deadline)

    if (-not $completionWindow) { throw '等待套版完成页超时。' }
    Show-AutomationWindow -Window $completionWindow
    Write-ControlSnapshot -Window $completionWindow -Label '完成页'
    Capture-Desktop -Name '04-完成页.png'
    Invoke-Control -Control (Get-Control -Root $completionWindow -Name '确定' -ControlType ([System.Windows.Automation.ControlType]::Button))
    if (-not $process.WaitForExit(15000)) { throw '关闭完成页后工具未退出。' }
    if ($process.ExitCode -ne 0) { throw "工具退出码异常：$($process.ExitCode)" }

    $newTaskDirectories = @(
        Get-ChildItem -LiteralPath $OutputRoot -Directory -ErrorAction SilentlyContinue |
            Where-Object { $existingTaskDirectoryPaths -notcontains $_.FullName }
    )
    if ($newTaskDirectories.Count -ne 1) {
        throw "本次应只生成 1 个任务文件夹，实际生成 $($newTaskDirectories.Count) 个。"
    }
    $taskDirectory = $newTaskDirectories[0]
    $jpgDirectory = Join-Path $taskDirectory.FullName 'JPG成品'
    $psdDirectory = Join-Path $taskDirectory.FullName 'PSD源文件'
    $recordDirectory = Join-Path $taskDirectory.FullName '任务记录'
    foreach ($requiredDirectory in @($jpgDirectory, $psdDirectory, $recordDirectory)) {
        if (-not (Test-Path -LiteralPath $requiredDirectory -PathType Container)) {
            throw "缺少任务输出目录：$requiredDirectory"
        }
    }

    $resultReport = Join-Path $recordDirectory '生成结果.csv'
    foreach ($requiredRecord in @('任务信息.txt', '任务日志.txt', 'data.csv', 'data_全部记录.csv', '异常记录.csv', '生成结果.csv')) {
        $requiredRecordPath = Join-Path $recordDirectory $requiredRecord
        if (-not (Test-Path -LiteralPath $requiredRecordPath -PathType Leaf)) {
            throw "缺少任务记录文件：$requiredRecordPath"
        }
    }

    $resultRows = @(Import-Csv -LiteralPath $resultReport -Encoding UTF8)
    if ($resultRows.Count -eq 0) { throw '生成结果.csv 没有商品结果。' }
    if ($UseSingleProduct -and $resultRows.Count -ne 1) {
        throw "单商品回归应产生 1 条结果，实际为 $($resultRows.Count) 条。"
    }
    if ($ProductCount -gt 0 -and $resultRows.Count -ne $ProductCount) {
        throw "指定商品回归应产生 $ProductCount 条结果，实际为 $($resultRows.Count) 条。"
    }

    $blockingStatuses = @('处理失败', '模板错误', '数据需核对', '缺图', '字段为空')
    $blockingRows = @($resultRows | Where-Object {
        $row = $_
        ($blockingStatuses -contains [string]$row.状态) -or
        ([string]$row.severity -eq 'E') -or
        @([string]$row.错误码 -split ';' | Where-Object { $_ -like 'E_*' }).Count -gt 0
    })
    if ($blockingRows.Count -gt 0) {
        foreach ($row in $blockingRows) {
            if (-not [string]::IsNullOrWhiteSpace([string]$row.输出文件) -or -not [string]::IsNullOrWhiteSpace([string]$row.输出PSD)) {
                throw "错误行不应导出 JPG 或 PSD：$($row.商品文件名) [$($row.错误码)]"
            }
        }
        $details = ($blockingRows | ForEach-Object { "$($_.商品文件名)：$($_.状态) [$($_.错误码)]" }) -join '；'
        throw "结果报告存在阻断错误：$details"
    }

    foreach ($row in $resultRows) {
        foreach ($outputField in @('输出文件', '输出PSD')) {
            $outputPath = [string]$row.$outputField
            if ([string]::IsNullOrWhiteSpace($outputPath) -or -not (Test-Path -LiteralPath $outputPath -PathType Leaf)) {
                throw "商品 $($row.商品文件名) 的$outputField不存在：$outputPath"
            }
        }
    }

    # Approved profiles isolate each variant below JPG成品/PSD源文件, while
    # legacy JD keeps the files directly under those folders. Count recursively
    # so the output gate validates both layouts without weakening file checks.
    $jpgFiles = @(Get-ChildItem -LiteralPath $jpgDirectory -Filter '*.jpg' -File -Recurse -ErrorAction SilentlyContinue)
    $psdFiles = @(Get-ChildItem -LiteralPath $psdDirectory -Filter '*.psd' -File -Recurse -ErrorAction SilentlyContinue)
    if ($jpgFiles.Count -ne $resultRows.Count -or $psdFiles.Count -ne $resultRows.Count) {
        throw "成品数与结果报告不一致：结果=$($resultRows.Count)，JPG=$($jpgFiles.Count)，PSD=$($psdFiles.Count)"
    }

    foreach ($jpg in $jpgFiles) {
        $image = [System.Drawing.Image]::FromFile($jpg.FullName)
        try {
            if ($image.Width -ne $ExpectedJpgWidth -or $image.Height -ne $ExpectedJpgHeight) {
                throw "JPG 尺寸异常：$($jpg.FullName)，$($image.Width)x$($image.Height)，预期 $($ExpectedJpgWidth)x$($ExpectedJpgHeight)"
            }
        } finally {
            $image.Dispose()
        }
    }
    foreach ($psd in $psdFiles) {
        $stream = [System.IO.File]::OpenRead($psd.FullName)
        try {
            $header = New-Object byte[] 4
            if ($stream.Read($header, 0, 4) -ne 4 -or [System.Text.Encoding]::ASCII.GetString($header) -ne '8BPS') {
                throw "PSD 文件头无效：$($psd.FullName)"
            }
        } finally {
            $stream.Dispose()
        }
    }

    $globalStatusPath = Join-Path $env:LOCALAPPDATA '电商主图套版工具\任务记录\status.json'
    if (-not (Test-Path -LiteralPath $globalStatusPath -PathType Leaf)) {
        throw "缺少工具最终状态：$globalStatusPath"
    }
    $globalStatus = Get-Content -LiteralPath $globalStatusPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($globalStatus.status -notin @('success', 'needs_review') -or [int]$globalStatus.exitCode -ne 0) {
        throw "工具最终状态异常：status=$($globalStatus.status)，exitCode=$($globalStatus.exitCode)"
    }

    Copy-Item -LiteralPath $jpgFiles[0].FullName -Destination (Join-Path $ArtifactDir '05-导出成品.jpg') -Force
    $warningCodes = @(
        $resultRows |
            ForEach-Object { [string]$_.错误码 -split ';' } |
            Where-Object { $_ -like 'W_*' } |
            Sort-Object -Unique
    )
    $warningText = if ($warningCodes.Count -gt 0) { $warningCodes -join ',' } else { 'none' }
    if ($RequireTextOverflow -and $warningCodes -notcontains 'W_TEXT_OVERFLOW') {
        throw '回归夹具应触发 W_TEXT_OVERFLOW，且该警告必须不阻断 JPG/PSD 导出。'
    }
    Write-SmokeLog "PASS：结果=$($resultRows.Count)，JPG=$($jpgFiles.Count)，PSD=$($psdFiles.Count)，status=$($globalStatus.status)，warnings=$warningText。"
    Set-Content -LiteralPath (Join-Path $ArtifactDir 'PASS.txt') -Value "Results=$($resultRows.Count)`r`nJPG=$($jpgFiles.Count)`r`nPSD=$($psdFiles.Count)`r`nStatus=$($globalStatus.status)`r`nWarnings=$warningText" -Encoding UTF8
} catch {
    Write-SmokeLog "FAIL：$($_.Exception.Message)"
    try { Capture-Desktop -Name 'FAIL.png' } catch {}
    if ($process -and -not $process.HasExited) {
        try { Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue } catch {}
    }
    Set-Content -LiteralPath (Join-Path $ArtifactDir 'FAIL.txt') -Value $_.Exception.ToString() -Encoding UTF8
    exit 1
}
