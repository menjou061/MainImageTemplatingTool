$signature = @'
using System;
using System.Runtime.InteropServices;

public static class MainImageConsoleWindow
{
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr handle, int command);
}
'@

Add-Type -TypeDefinition $signature

$processIds = New-Object System.Collections.Generic.List[int]
$currentId = $PID
for ($depth = 0; $depth -lt 5 -and $currentId -gt 0; $depth++) {
    $processIds.Add([int]$currentId)
    $process = Get-CimInstance Win32_Process -Filter "ProcessId = $currentId" -ErrorAction SilentlyContinue
    if (-not $process -or [int]$process.ParentProcessId -eq $currentId) {
        break
    }
    $currentId = [int]$process.ParentProcessId
}

foreach ($processId in $processIds) {
    $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
    if ($process -and $process.ProcessName -ieq 'cmd') {
        $handle = [IntPtr]$process.MainWindowHandle
        if ($handle -ne [IntPtr]::Zero) {
            [void][MainImageConsoleWindow]::ShowWindow($handle, 0)
        }
    }
}

$consoleHandle = [MainImageConsoleWindow]::GetConsoleWindow()
if ($consoleHandle -ne [IntPtr]::Zero) {
    [void][MainImageConsoleWindow]::ShowWindow($consoleHandle, 0)
}
