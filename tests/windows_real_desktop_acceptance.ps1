param(
    [Parameter(Mandatory = $true)][string]$ToolZipPath,
    [Parameter(Mandatory = $true)][string]$ExcelPath,
    [Parameter(Mandatory = $true)][string]$PsdPath,
    [Parameter(Mandatory = $true)][string]$WorkRoot,
    [int]$ExpectedWidth = 750,
    [int]$ExpectedHeight = 1000,
    [switch]$UseSingleProduct
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

function Assert-File {
    param([string]$Path, [string]$Label)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "$Label 不存在：$Path"
    }
}

function Expand-Release {
    param([string]$ArchivePath, [string]$DestinationPath)
    New-Item -Path $DestinationPath -ItemType Directory -Force | Out-Null
    $tar = Get-Command tar.exe -ErrorAction SilentlyContinue
    if ($tar) {
        & $tar.Source -xf $ArchivePath -C $DestinationPath
        if ($LASTEXITCODE -eq 0) { return }
        throw "tar.exe 解压失败，退出码：$LASTEXITCODE"
    }
    Expand-Archive -LiteralPath $ArchivePath -DestinationPath $DestinationPath -Force
}

Assert-File -Path $ToolZipPath -Label 'Release ZIP'
Assert-File -Path $ExcelPath -Label 'real channel workbook'
Assert-File -Path $PsdPath -Label 'real channel PSD'
if (Test-Path -LiteralPath $WorkRoot) {
    throw "work directory must be new: $WorkRoot"
}

$photoshop = @(Get-Process -Name Photoshop -ErrorAction SilentlyContinue)
if ($photoshop.Count -eq 0) {
    throw 'Photoshop must be started and logged in.'
}

$extractRoot = Join-Path $WorkRoot 'extracted'
$outputRoot = Join-Path $WorkRoot 'output'
$artifactRoot = Join-Path $WorkRoot 'artifacts'
New-Item -Path $WorkRoot -ItemType Directory | Out-Null
Expand-Release -ArchivePath $ToolZipPath -DestinationPath $extractRoot

$entry = Get-ChildItem -LiteralPath $extractRoot -Filter '*.cmd' -File -Recurse | Where-Object { $_.Name -ne 'L0_Start.cmd' } | Select-Object -First 1
if (-not $entry) { throw 'Release ZIP entrypoint is missing.' }
$toolRoot = $entry.DirectoryName

$manifest = [ordered]@{
    source = 'Windows Desktop real channel files'
    excel = [ordered]@{ path = $ExcelPath; sha256 = (Get-FileHash -LiteralPath $ExcelPath -Algorithm SHA256).Hash.ToLowerInvariant(); bytes = (Get-Item -LiteralPath $ExcelPath).Length }
    psd = [ordered]@{ path = $PsdPath; sha256 = (Get-FileHash -LiteralPath $PsdPath -Algorithm SHA256).Hash.ToLowerInvariant(); bytes = (Get-Item -LiteralPath $PsdPath).Length }
    expected_size = "$ExpectedWidth`x$ExpectedHeight"
    use_single_product = [bool]$UseSingleProduct
    started_at = (Get-Date).ToString('o')
}
New-Item -Path $artifactRoot -ItemType Directory -Force | Out-Null
$manifest | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $artifactRoot 'real_input_manifest.json') -Encoding UTF8

$uiSmoke = Join-Path $PSScriptRoot 'windows_ui_smoke.ps1'
Assert-File -Path $uiSmoke -Label 'Windows UI smoke script'
& $uiSmoke -ToolRoot $toolRoot -ExcelPath $ExcelPath -PsdPath $PsdPath -OutputRoot $outputRoot -ArtifactDir $artifactRoot -ExpectedWidth $ExpectedWidth -ExpectedHeight $ExpectedHeight -UseSingleProduct:$UseSingleProduct -UseHygieneTmall
if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath (Join-Path $artifactRoot 'PASS.txt') -PathType Leaf)) {
    throw "real desktop acceptance failed; inspect: $artifactRoot"
}

Write-Host 'PASS: Windows desktop real channel Excel + PSD exported through Photoshop.' -ForegroundColor Green
Write-Host "input manifest: $(Join-Path $artifactRoot 'real_input_manifest.json')"
Write-Host "acceptance root: $WorkRoot"
