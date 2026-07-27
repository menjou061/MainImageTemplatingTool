param(
    [Parameter(Mandatory = $true)][string]$ZipPath,
    [Parameter(Mandatory = $true)][string]$PsdPath,
    [Parameter(Mandatory = $true)][string]$ProductImagePath,
    [Parameter(Mandatory = $true)][string]$WorkRoot
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

function Assert-PathIsFile {
    param([string]$Path, [string]$Purpose)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "$Purpose 不存在：$Path"
    }
}

Assert-PathIsFile -Path $ZipPath -Purpose 'Release ZIP'
Assert-PathIsFile -Path $PsdPath -Purpose '验收 PSD'
Assert-PathIsFile -Path $ProductImagePath -Purpose '验收商品图'
if (Test-Path -LiteralPath $WorkRoot) {
    throw "验收目录必须是尚不存在的新目录，避免旧产物造成误判：$WorkRoot"
}

$photoshopProcesses = @(Get-Process -Name Photoshop -ErrorAction SilentlyContinue)
if ($photoshopProcesses.Count -eq 0) {
    throw '请先启动并登录 Photoshop、进入首页，再运行 Release 验收。'
}

New-Item -Path $WorkRoot -ItemType Directory | Out-Null
$extractRoot = Join-Path $WorkRoot 'extracted'
$outputRoot = Join-Path $WorkRoot 'output'
$artifactRoot = Join-Path $WorkRoot 'artifacts'
$workbookPath = Join-Path $WorkRoot 'release_e2e.xlsx'

Expand-Archive -LiteralPath $ZipPath -DestinationPath $extractRoot -Force
$allItems = @(Get-ChildItem -LiteralPath $extractRoot -Force -Recurse)
$entries = @($allItems | Where-Object { -not $_.PSIsContainer })
if ($entries.Count -lt 1) { throw 'Release ZIP 解压后没有文件。' }
if (@($allItems | Where-Object { $_.Name -eq '__pycache__' -or $_.Extension -eq '.pyc' }).Count -gt 0) {
    throw 'Release ZIP 包含 __pycache__ 或 .pyc。'
}

$entry = $entries | Where-Object { $_.Name -eq '开始套版.cmd' } | Select-Object -First 1
if (-not $entry) { throw 'Release ZIP 缺少开始套版.cmd。' }
$toolRoot = $entry.DirectoryName
foreach ($relativePath in @(
    'L0_Start.cmd',
    '_internal\L0_Run.bat',
    '_internal\L0_Run.ps1',
    '_internal\runtime\python\python.exe',
    '示例文件\表格案例_618正式主图.xlsx'
)) {
    if (-not (Test-Path -LiteralPath (Join-Path $toolRoot $relativePath) -PathType Leaf)) {
        throw "Release ZIP 缺少交付文件：$relativePath"
    }
}

foreach ($launcherPath in @(
    $entry.FullName,
    (Join-Path $toolRoot 'L0_Start.cmd'),
    (Join-Path $toolRoot '_internal\L0_Run.bat')
)) {
    $bytes = [System.IO.File]::ReadAllBytes($launcherPath)
    if (@($bytes | Where-Object { $_ -gt 127 }).Count -gt 0) {
        throw "启动脚本必须保持 ASCII：$launcherPath"
    }
    $text = [System.Text.Encoding]::ASCII.GetString($bytes)
    if ($text -match '(?<!\r)\n') {
        throw "启动脚本必须使用 CRLF 换行：$launcherPath"
    }
}

$python = Join-Path $toolRoot '_internal\runtime\python\python.exe'
$runtimeProbe = & $python -B -c 'import openpyxl, et_xmlfile, sys; print(sys.version.split()[0]); print(openpyxl.__version__); print(et_xmlfile.__version__)'
if ($LASTEXITCODE -ne 0 -or $runtimeProbe.Count -ne 3 -or
    $runtimeProbe[0] -ne '3.11.9' -or
    $runtimeProbe[1] -ne '3.1.5' -or
    $runtimeProbe[2] -ne '2.0.0') {
    throw '随包 Python 运行时检测失败。'
}

$fixtureScript = Join-Path $PSScriptRoot 'create_windows_e2e_fixture.py'
Assert-PathIsFile -Path $fixtureScript -Purpose 'E2E 表格生成脚本'
& $python -B $fixtureScript --output $workbookPath --product-image $ProductImagePath
if ($LASTEXITCODE -ne 0) { throw 'E2E 表格生成失败。' }

$uiSmoke = Join-Path $PSScriptRoot 'windows_ui_smoke.ps1'
Assert-PathIsFile -Path $uiSmoke -Purpose 'Windows UI 回归脚本'
& $uiSmoke -ToolRoot $toolRoot -ExcelPath $workbookPath -PsdPath $PsdPath -OutputRoot $outputRoot -ArtifactDir $artifactRoot -UseSingleProduct
if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath (Join-Path $artifactRoot 'PASS.txt') -PathType Leaf)) {
    throw "Release E2E 未通过，请查看：$artifactRoot"
}

Write-Host 'PASS：Release ZIP 已完成全新解压、运行时检查和真实单商品导出。'
Write-Host "验收目录：$WorkRoot"
Write-Host "运行时：Python $($runtimeProbe[0]) / openpyxl $($runtimeProbe[1]) / et-xmlfile $($runtimeProbe[2])"
