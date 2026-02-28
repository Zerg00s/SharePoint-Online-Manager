# Requires: Azure CLI logged in (az login), Artifact Signing Client Tools installed
# Install: winget install -e --id Microsoft.Azure.ArtifactSigningClientTools

param(
    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"

$projectDir = "$PSScriptRoot\.."
$outputExe  = "$projectDir\bin\Release\net8.0-windows\SharePoint-Online-Manager.exe"
$metadataJson = "$PSScriptRoot\signing-metadata.json"

# Step 1: Build Release
if (-not $SkipBuild) {
    Write-Host "Building Release..." -ForegroundColor Cyan
    dotnet build $projectDir -c Release
    if ($LASTEXITCODE -ne 0) { throw "Build failed" }
}

# Step 2: Verify exe exists
if (-not (Test-Path $outputExe)) { throw "EXE not found: $outputExe" }

# Step 3: Sign with Azure Artifact Signing
Write-Host "Signing $outputExe ..." -ForegroundColor Cyan

# SignTool path (installed with Artifact Signing Client Tools or Windows SDK)
$signToolPaths = @(
    "${env:ProgramFiles(x86)}\Windows Kits\10\bin\*\x64\signtool.exe",
    "${env:ProgramFiles}\Azure Code Signing\signtool.exe"
)
$signTool = $signToolPaths | ForEach-Object { Get-Item $_ -ErrorAction SilentlyContinue } |
            Sort-Object FullName -Descending | Select-Object -First 1

if (-not $signTool) { throw "SignTool.exe not found. Install Windows SDK or Artifact Signing Client Tools." }

# Dlib path (installed with Artifact Signing Client Tools or NuGet)
$dlibPaths = @(
    "C:\trash\SigningTools\Microsoft.Trusted.Signing.Client.*\bin\x64\Azure.CodeSigning.Dlib.dll",
    "${env:ProgramFiles}\Azure Code Signing\Azure.CodeSigning.Dlib.dll",
    "$PSScriptRoot\packages\Microsoft.ArtifactSigning.Client\*\bin\x64\Azure.CodeSigning.Dlib.dll"
)
$dlib = $dlibPaths | ForEach-Object { Get-Item $_ -ErrorAction SilentlyContinue } |
        Sort-Object FullName -Descending | Select-Object -First 1

if (-not $dlib) { throw "Azure.CodeSigning.Dlib.dll not found. Install: winget install -e --id Microsoft.Azure.ArtifactSigningClientTools" }

& $signTool.FullName sign /v /debug /fd SHA256 `
    /tr "http://timestamp.acs.microsoft.com" /td SHA256 `
    /dlib $dlib.FullName `
    /dmdf $metadataJson `
    $outputExe

if ($LASTEXITCODE -ne 0) { throw "Signing failed" }

# Step 4: Verify
Write-Host "Verifying signature..." -ForegroundColor Cyan
& $signTool.FullName verify /pa /v $outputExe

Write-Host "`nDone! Signed exe: $outputExe" -ForegroundColor Green
