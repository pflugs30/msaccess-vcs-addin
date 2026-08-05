#
# Install-BootstrapAddin.ps1
#
# Installs a stable PUBLIC msaccess-vcs add-in release to bootstrap a CI build
# of this repository's own source. The bootstrap add-in provides the "Build
# From Source" engine (btnBuild) that compiles Version Control.accda.src into a
# working database.
#
# Ported (trimmed) from msaccess-vcs-build/scripts/Install-msaccess-vcs.ps1 and
# Set-TrustedLocation.ps1 so this repository's release pipeline is self-contained
# and does not depend on the fork build actions.
#
param(
    # Public add-in release API URL. Defaults to the pinned v5.0.1 public release.
    [string]$VcsUrl = "https://api.github.com/repos/joyfullservice/msaccess-vcs-addin/releases/tags/v5.0.1",

    # Optional folder to add as an Access Trusted Location (AllowSubfolders=1).
    # Pass the workspace root so both the source tree and the build output are
    # trusted and Access does not block on a macro-security prompt.
    [string]$TrustSourceDir = ""
)

$ErrorActionPreference = "Stop"

function Set-AccessTrustedLocation {
    param([string]$Name, [string]$Path)
    $regPath = "HKCU:\Software\Microsoft\Office\16.0\Access\Security\Trusted Locations\$Name"
    if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
    New-ItemProperty -Path $regPath -Name "Path" -Value $Path -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $regPath -Name "AllowSubfolders" -Value 1 -PropertyType DWord -Force | Out-Null
    Write-Host "Trusted location set: $Name -> $Path"
}

$headers = @{ "User-Agent" = "PowerShell" }
$githubToken = $env:GH_TOKEN
if ([string]::IsNullOrEmpty($githubToken)) { $githubToken = $env:GITHUB_TOKEN }
if (-not [string]::IsNullOrEmpty($githubToken)) {
    $headers["Authorization"] = "Bearer $githubToken"
}

Write-Host "Bootstrap add-in release: $VcsUrl"
$release = Invoke-RestMethod -Uri $VcsUrl -Headers $headers

$asset = $release.assets | Where-Object { $_.name -like "Version*.zip" } | Select-Object -First 1
if ($null -eq $asset) {
    Write-Error "No 'Version*.zip' asset found on the bootstrap release: $VcsUrl"
    exit 1
}

$zipFile = Join-Path $env:RUNNER_TEMP "bootstrap-msaccess-vcs.zip"
Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $zipFile
Write-Host "Downloaded bootstrap zip: $($asset.name)"

# Install to the default add-in location so the build script's default add-in
# path (%APPDATA%\MSAccessVCS\Version Control) resolves without extra wiring.
$addInFolder = Join-Path $env:APPDATA "MSAccessVCS"
Expand-Archive -Path $zipFile -DestinationPath $addInFolder -Force

$addInPath = Join-Path $addInFolder "Version Control.accda"
if (-not (Test-Path $addInPath)) {
    Write-Error "Bootstrap add-in not found after extraction: $addInPath"
    exit 1
}
Write-Host "Bootstrap add-in installed: $addInPath"

Set-AccessTrustedLocation -Name "VCS-add-in-folder" -Path $addInFolder
if (-not [string]::IsNullOrWhiteSpace($TrustSourceDir)) {
    Set-AccessTrustedLocation -Name "VCS-build-folder" -Path $TrustSourceDir
}
