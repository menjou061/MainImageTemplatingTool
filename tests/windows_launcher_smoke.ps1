param(
    [Parameter(Mandatory = $true)][string]$ToolRoot,
    [Parameter(Mandatory = $true)][string]$ArtifactDir
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

New-Item -Path $ArtifactDir -ItemType Directory -Force | Out-Null
$logPath = Join-Path $ArtifactDir 'launcher-smoke.log'
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
        [int]$TimeoutSeconds = 45
    )
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        $window = Get-DesktopWindow -Title $Title
        if ($window) { return $window }
        if ($process -and $process.HasExited) {
            throw "启动入口提前退出，退出码：$($process.ExitCode)"
        }
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
    return $Root.FindFirst(
        [System.Windows.Automation.TreeScope]::Descendants,
        (New-Object System.Windows.Automation.AndCondition($conditions))
    )
}

try {
    $entry = Join-Path $ToolRoot '开始套版.cmd'
    if (-not (Test-Path -LiteralPath $entry -PathType Leaf)) {
        throw "未找到用户启动入口：$entry"
    }
    $mainScript = Join-Path $ToolRoot '_internal\L0_Run.ps1'
    if (-not (Test-Path -LiteralPath $mainScript -PathType Leaf)) {
        throw "未找到主程序：$mainScript"
    }
    $invalidPointExpressions = @(Select-String -LiteralPath $mainScript -Pattern 'New-Object\s+System\.Drawing\.Point\([^\r\n]*,\s*\(if\s*\(')
    if ($invalidPointExpressions.Count -gt 0) {
        throw "主程序包含无法执行的 Point(..., (if ...)) 表达式，共 $($invalidPointExpressions.Count) 处。"
    }

    Write-SmokeLog "从用户入口启动：$entry"
    $arguments = '/d /c call "{0}"' -f $entry
    $process = Start-Process -FilePath (Join-Path $env:SystemRoot 'System32\cmd.exe') -ArgumentList $arguments -WindowStyle Normal -PassThru

    $channelWindow = Wait-DesktopWindow -Title '选择品类和渠道'
    $nextButton = Get-Control -Root $channelWindow -Name '下一步' -ControlType ([System.Windows.Automation.ControlType]::Button)
    if (-not $nextButton) { throw '品类渠道选择页缺少下一步按钮。' }
    $nextButton.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()

    $mainWindow = Wait-DesktopWindow -Title '电商主图套版工具 1.2'
    $startupHint = Get-Control -Root $mainWindow -Name '请先启动并登录 Photoshop，进入首页后再选择商品表格和 PSD 模板。' -ControlType ([System.Windows.Automation.ControlType]::Text)
    if (-not $startupHint) { throw '初始页缺少 Photoshop 启动提醒。' }

    $cancelButton = Get-Control -Root $mainWindow -Name '取消' -ControlType ([System.Windows.Automation.ControlType]::Button)
    if (-not $cancelButton) { throw '初始页缺少取消按钮。' }
    $cancelButton.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()

    if (-not $process.WaitForExit(15000)) { throw '取消后启动进程未退出。' }
    if ($process.ExitCode -ne 0) { throw "取消后退出码异常：$($process.ExitCode)" }

    Write-SmokeLog 'PASS：中文入口成功打开主窗口，取消后正常退出。'
    Set-Content -LiteralPath (Join-Path $ArtifactDir 'PASS.txt') -Value 'Launcher opened the main window and exited cleanly.' -Encoding UTF8
} catch {
    Write-SmokeLog "FAIL：$($_.Exception.Message)"
    if ($process -and -not $process.HasExited) {
        try { & taskkill.exe /PID $process.Id /T /F | Out-Null } catch {}
    }
    Set-Content -LiteralPath (Join-Path $ArtifactDir 'FAIL.txt') -Value $_.Exception.ToString() -Encoding UTF8
    exit 1
}
