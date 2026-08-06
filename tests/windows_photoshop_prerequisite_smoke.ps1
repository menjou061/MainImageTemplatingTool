param(
    [Parameter(Mandatory = $true)][string]$ToolRoot,
    [Parameter(Mandatory = $true)][string]$ArtifactDir,
    [Parameter(Mandatory = $true)][string]$OutputRoot,
    [ValidateSet('NotRunning', 'NotReady')][string]$ExpectedState = 'NotRunning'
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class NativeWindow {
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr window);
}
'@

New-Item -Path $ArtifactDir -ItemType Directory -Force | Out-Null
$logPath = Join-Path $ArtifactDir 'photoshop-prerequisite-smoke.log'
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
    try {
        $windowPattern = $Window.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern)
        if ($windowPattern.Current.WindowVisualState -eq [System.Windows.Automation.WindowVisualState]::Minimized) {
            $windowPattern.SetWindowVisualState([System.Windows.Automation.WindowVisualState]::Normal)
        }
    } catch {
    }
    [void][NativeWindow]::SetForegroundWindow([IntPtr]$Window.Current.NativeWindowHandle)
    Start-Sleep -Milliseconds 500
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

try {
    $photoshopProcesses = @(Get-Process -Name 'Photoshop' -ErrorAction SilentlyContinue)
    if ($ExpectedState -eq 'NotRunning' -and $photoshopProcesses.Count -gt 0) {
        throw '前置条件不满足：Photoshop 仍在运行。'
    }
    if ($ExpectedState -eq 'NotReady' -and $photoshopProcesses.Count -eq 0) {
        throw '前置条件不满足：没有可用于未就绪检查的 Photoshop 进程。'
    }

    Write-SmokeLog '开始 Photoshop 未启动流程回归。'
    $entry = Join-Path $ToolRoot '开始套版.cmd'
    if (-not (Test-Path -LiteralPath $entry -PathType Leaf)) {
        throw "未找到用户启动入口：$entry"
    }
    $arguments = '/d /c call "{0}"' -f $entry
    $process = Start-Process -FilePath (Join-Path $env:SystemRoot 'System32\cmd.exe') -ArgumentList $arguments -WindowStyle Normal -PassThru

    $channelWindow = Wait-DesktopWindow -Title '选择品类和渠道' -TimeoutSeconds 45
    Invoke-Control -Control (Get-Control -Root $channelWindow -Name '下一步' -ControlType ([System.Windows.Automation.ControlType]::Button))

    $mainWindow = Wait-DesktopWindow -Title '电商主图套版工具 1.4' -TimeoutSeconds 45
    $startupHint = Get-Control -Root $mainWindow -Name '请先启动并登录 Photoshop，进入首页后再选择商品表格和 PSD 模板。' -ControlType ([System.Windows.Automation.ControlType]::Text)
    if (-not $startupHint) { throw '初始页缺少 Photoshop 启动提醒。' }

    $edits = @($mainWindow.FindAll(
        [System.Windows.Automation.TreeScope]::Descendants,
        (New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
            [System.Windows.Automation.ControlType]::Edit
        ))) | Sort-Object { $_.Current.BoundingRectangle.Y })
    if ($edits.Count -lt 3) { throw "输入框数量异常：$($edits.Count)" }
    $edits[2].GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).SetValue($OutputRoot)

    Invoke-Control -Control (Get-Control -Root $mainWindow -Name '开始生成' -ControlType ([System.Windows.Automation.ControlType]::Button))
    Start-Sleep -Seconds 1

    $mainWindow = Get-DesktopWindow -Title '电商主图套版工具 1.4'
    if (-not $mainWindow) { throw '点击开始生成后设置页被关闭。' }
    if ($process.HasExited) { throw "工具意外退出，退出码：$($process.ExitCode)" }
    $expectedMessage = if ($ExpectedState -eq 'NotRunning') {
        '请先启动并登录 Photoshop，进入首页后再点击“开始生成”。'
    } else {
        'Photoshop 正在启动或尚未显示首页。请处理登录、授权或弹窗，进入首页后再试。'
    }
    $inlineHint = Get-Control -Root $mainWindow -Name $expectedMessage -ControlType ([System.Windows.Automation.ControlType]::Text)
    if (-not $inlineHint) { throw '没有显示预期的 Photoshop 行内提醒。' }
    if (Test-Path -LiteralPath $OutputRoot) { throw 'Photoshop 前置检查失败后仍创建了输出目录。' }

    Show-AutomationWindow -Window $mainWindow
    Capture-Desktop -Name '01-Photoshop未启动提醒.png'
    Invoke-Control -Control (Get-Control -Root $mainWindow -Name '取消' -ControlType ([System.Windows.Automation.ControlType]::Button))
    if (-not $process.WaitForExit(15000)) { throw '取消后工具未退出。' }
    if ($process.ExitCode -ne 0) { throw "工具退出码异常：$($process.ExitCode)" }

    Write-SmokeLog "PASS：$ExpectedState 设置页保留，行内提醒已显示，未创建输出目录。"
    Set-Content -LiteralPath (Join-Path $ArtifactDir 'PASS.txt') -Value "Photoshop prerequisite reminder verified: $ExpectedState" -Encoding UTF8
} catch {
    Write-SmokeLog "FAIL：$($_.Exception.Message)"
    try { Capture-Desktop -Name 'FAIL.png' } catch {}
    if ($process -and -not $process.HasExited) {
        try { Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue } catch {}
    }
    Set-Content -LiteralPath (Join-Path $ArtifactDir 'FAIL.txt') -Value $_.Exception.ToString() -Encoding UTF8
    exit 1
}
