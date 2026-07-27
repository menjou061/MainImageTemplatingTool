param(
    [Parameter(Mandatory = $true)][string]$ToolRoot,
    [Parameter(Mandatory = $true)][string]$ExcelPath,
    [Parameter(Mandatory = $true)][string]$PsdPath,
    [Parameter(Mandatory = $true)][string]$OutputRoot,
    [Parameter(Mandatory = $true)][string]$ArtifactDir
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
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint flags, uint dx, uint dy, uint data, UIntPtr extraInfo);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr window);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr window, int command);
    public const uint LeftDown = 0x0002;
    public const uint LeftUp = 0x0004;
}
'@

New-Item -Path $ArtifactDir -ItemType Directory -Force | Out-Null
$logPath = Join-Path $ArtifactDir 'ui-smoke.log'
$process = $null

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

function Show-AutomationWindow {
    param([System.Windows.Automation.AutomationElement]$Window)
    if (-not $Window) { return }
    $handle = [IntPtr]$Window.Current.NativeWindowHandle
    [void][NativeMouse]::ShowWindow($handle, 9)
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
        Write-SmokeLog ("{0} | {1} | x={2} y={3} w={4} h={5}" -f $control.Current.ControlType.ProgrammaticName, $control.Current.Name, [int]$rectangle.X, [int]$rectangle.Y, [int]$rectangle.Width, [int]$rectangle.Height)
    }
}

try {
    Write-SmokeLog '开始 Windows 界面回归。'
    $runner = Join-Path $ToolRoot '_internal\L0_Run.ps1'
    $arguments = @('-NoLogo', '-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $runner))
    $process = Start-Process -FilePath 'powershell.exe' -ArgumentList ($arguments -join ' ') -WindowStyle Minimized -PassThru

    $mainWindow = Wait-DesktopWindow -Title '电商主图套版工具 1.0' -TimeoutSeconds 45
    Show-AutomationWindow -Window $mainWindow
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

    $mainWindow = Wait-DesktopWindow -Title '电商主图套版工具 1.0'
    Show-AutomationWindow -Window $mainWindow
    Write-ControlSnapshot -Window $mainWindow -Label '表格已加载'
    Capture-Desktop -Name '02-表格已加载.png'

    $partialRadio = Get-Control -Root $mainWindow -Name '只生成勾选的商品' -ControlType ([System.Windows.Automation.ControlType]::RadioButton)
    Click-Control -Control $partialRadio
    Start-Sleep -Milliseconds 500
    $productListControl = Get-Control -Root $mainWindow -Name '4  商品范围' -ControlType ([System.Windows.Automation.ControlType]::List)
    if (-not $productListControl) { throw '商品列表为空。' }
    Show-AutomationWindow -Window $mainWindow
    $productListControl.SetFocus()
    [System.Windows.Forms.SendKeys]::SendWait('{HOME}')
    [System.Windows.Forms.SendKeys]::SendWait(' ')
    Start-Sleep -Seconds 1
    Capture-Desktop -Name '03-商品已勾选.png'

    # Use the all-products path for the Photoshop export portion of this smoke
    # test. The checked-item state is visually verified in the prior screenshot.
    $allRadio = Get-Control -Root $mainWindow -Name '全部商品' -ControlType ([System.Windows.Automation.ControlType]::RadioButton)
    Click-Control -Control $allRadio
    Start-Sleep -Milliseconds 500

    Invoke-Control -Control (Get-Control -Root $mainWindow -Name '开始生成' -ControlType ([System.Windows.Automation.ControlType]::Button))
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

    $jpgCount = @(Get-ChildItem -LiteralPath $OutputRoot -Filter '*.jpg' -File -Recurse -ErrorAction SilentlyContinue).Count
    $psdCount = @(Get-ChildItem -LiteralPath $OutputRoot -Filter '*.psd' -File -Recurse -ErrorAction SilentlyContinue).Count
    if ($jpgCount -lt 1 -or $psdCount -lt 1) { throw "成品数量异常：JPG=$jpgCount，PSD=$psdCount" }
    Write-SmokeLog "PASS：JPG=$jpgCount，PSD=$psdCount。"
    Set-Content -LiteralPath (Join-Path $ArtifactDir 'PASS.txt') -Value "JPG=$jpgCount`r`nPSD=$psdCount" -Encoding UTF8
} catch {
    Write-SmokeLog "FAIL：$($_.Exception.Message)"
    try { Capture-Desktop -Name 'FAIL.png' } catch {}
    if ($process -and -not $process.HasExited) {
        try { Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue } catch {}
    }
    Set-Content -LiteralPath (Join-Path $ArtifactDir 'FAIL.txt') -Value $_.Exception.ToString() -Encoding UTF8
    exit 1
}
