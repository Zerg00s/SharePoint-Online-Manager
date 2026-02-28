# Requires: Azure CLI logged in (az login), Artifact Signing Client Tools installed
# Install: winget install -e --id Microsoft.Azure.ArtifactSigningClientTools
#
# Usage:
#   .\Sign-App.ps1 -Version 1.0.0                  # publish + sign
#   .\Sign-App.ps1 -Version 1.0.0 -Release          # publish + sign + push GitHub release
#   .\Sign-App.ps1 -Version 1.0.0 -SkipBuild        # sign only (already published)
#   .\Sign-App.ps1 -Version 1.0.0 -Release -Draft    # push as draft release

param(
    [Parameter(Mandatory=$true)]
    [string]$Version,
    [switch]$SkipBuild,
    [switch]$Release,
    [switch]$Draft
)

$ErrorActionPreference = "Stop"

$projectDir   = $PSScriptRoot
$publishDir   = "$projectDir\bin\Release\net8.0-windows\win-x64\publish"
$outputExe    = "$publishDir\SharePoint-Online-Manager.exe"
$metadataJson = "$PSScriptRoot\signing-metadata.json"
$tag          = "v$Version"

# Step 1: Publish self-contained single-file exe
if (-not $SkipBuild) {
    Write-Host "Publishing Release v$Version (self-contained, single file)..." -ForegroundColor Cyan
    dotnet publish $projectDir -c Release -r win-x64 --self-contained `
        -p:PublishSingleFile=true `
        -p:IncludeNativeLibrariesForSelfExtract=true `
        -p:Version=$Version
    if ($LASTEXITCODE -ne 0) { throw "Publish failed" }
}

# Step 2: Verify exe exists
if (-not (Test-Path $outputExe)) { throw "EXE not found: $outputExe" }

Write-Host "Published exe: $outputExe ($('{0:N1} MB' -f ((Get-Item $outputExe).Length / 1MB)))" -ForegroundColor Gray

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

# Step 4: Verify signature
Write-Host "Verifying signature..." -ForegroundColor Cyan
& $signTool.FullName verify /pa /v $outputExe

Write-Host "`nSigned exe: $outputExe" -ForegroundColor Green

# Step 5: Push GitHub release (if -Release flag)
if ($Release) {
    Write-Host "`nCreating GitHub release $tag ..." -ForegroundColor Cyan

    # Verify gh CLI is available
    if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
        throw "GitHub CLI (gh) not found. Install: winget install -e --id GitHub.cli"
    }

    $repoRoot = (Resolve-Path "$projectDir\..").Path
    Push-Location $repoRoot

    try {
        $ghArgs = @("release", "create", $tag, $outputExe, "--title", "SharePoint Online Manager $tag", "--generate-notes")

        if ($Draft) {
            $ghArgs += "--draft"
        }

        & gh @ghArgs
        if ($LASTEXITCODE -ne 0) { throw "GitHub release failed" }

        Write-Host "`nGitHub release $tag created!" -ForegroundColor Green
    }
    finally {
        Pop-Location
    }
}
