#
# Build-AddinFromSource.ps1
#
# Builds the msaccess-vcs add-in from its own text source (Version
# Control.accda.src) using an already-installed bootstrap add-in, driven through
# the Access COM automation model. Emits the built database path as the step
# output "builtPath".
#
# Ported (trimmed) from msaccess-vcs-build/scripts/Build-Accdb.ps1. The heavy
# hang-diagnostics from the reference script are intentionally omitted here; add
# them back if a headless run stalls on a modal Access dialog.
#
param(
    # Absolute path to the exported source folder (Version Control.accda.src).
    [Parameter(Mandatory = $true)][string]$SourceDir,

    # Working/output folder. The temp build database and the VCS build output
    # land here (kept separate from the repo tree to avoid clobbering the
    # committed Version Control.accda binary).
    [string]$OutputDir = ""
)

$ErrorActionPreference = "Stop"

if (-not [System.IO.Path]::IsPathRooted($SourceDir)) {
    $SourceDir = Join-Path (Get-Location) $SourceDir
}
if (-not (Test-Path $SourceDir)) {
    Write-Error "Source folder not found: $SourceDir"
    exit 1
}

if ([string]::IsNullOrWhiteSpace($OutputDir)) { $OutputDir = (Get-Location).Path }
if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }

# Resolve any relative build output against the output folder, not the repo root.
Set-Location $OutputDir

$tempFileName = "VcsBuildTempApp"
$accdbPath = Join-Path $OutputDir "$tempFileName.accdb"
if (Test-Path $accdbPath) { Remove-Item $accdbPath -Force }

# Default installed add-in path (matches Install-BootstrapAddin.ps1 target).
$addInProcessPath = Join-Path (Join-Path $env:APPDATA "MSAccessVCS") "Version Control"
if (-not (Test-Path "$addInProcessPath.accd[ae]")) {
    Write-Error "Bootstrap add-in not found: $addInProcessPath.accda (run Install-BootstrapAddin.ps1 first)"
    exit 1
}

Write-Host "Source:  $SourceDir"
Write-Host "Add-in:  $addInProcessPath"
Write-Host "Output:  $OutputDir"

$access = New-Object -ComObject Access.Application
$access.Visible = $true

# Track the specific MSACCESS process so we can wait for its file lock to clear
# after the build (Quit() returns before the process actually exits).
$accessProcessId = $null
try {
    $handle = [int]($access.hWndAccessApp())
    for ($i = 0; $i -lt 10 -and $null -eq $accessProcessId; $i++) {
        $proc = Get-Process -Name MSACCESS -ErrorAction SilentlyContinue |
            Where-Object { $_.MainWindowHandle -eq $handle } | Select-Object -First 1
        if ($proc) { $accessProcessId = $proc.Id; break }
        Start-Sleep -Milliseconds 200
    }
}
catch { $accessProcessId = $null }

$access.NewCurrentDatabase($accdbPath)

Write-Host "Starting VCS build from source..."
$access.Run("$addInProcessPath.SetInteractionMode", [ref] 1)
$null = $access.Run("$addInProcessPath.HandleRibbonCommand", [ref] "btnBuild", [ref] "$SourceDir")

# The build closes the temp app and opens the freshly built database. Poll for
# forms to settle across two passes, matching the reference implementation.
foreach ($pass in 1..2) {
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while (($access.Forms.Count -gt 0) -and ($sw.Elapsed.TotalSeconds -lt 180)) {
        Start-Sleep -Seconds 2
        Write-Host "." -NoNewline
    }
    $sw.Stop()
    Start-Sleep -Seconds 3
}
Write-Host ""

$builtFileName = $access.CurrentProject.Name
$builtFilePath = $access.CurrentProject.FullName

Start-Sleep -Seconds 1
$access.Quit(2)
[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($access)
Remove-Variable access
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

if ([string]::IsNullOrWhiteSpace($builtFileName) -or ($builtFileName -eq "$tempFileName.accdb")) {
    Write-Error "Build failed (built name: '$builtFileName', path: '$builtFilePath')"
    exit 1
}
Write-Host "Built: $builtFileName ($builtFilePath)"

# Wait for the tracked Access process to exit and release the file lock before
# the workflow packages the built file.
if ($null -ne $accessProcessId) {
    Write-Host "Waiting for Access to shut down..."
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ((Get-Process -Id $accessProcessId -ErrorAction SilentlyContinue) -and ($sw.Elapsed.TotalSeconds -lt 180)) {
        Start-Sleep -Milliseconds 500
    }
    $sw.Stop()
}

if (-not (Test-Path $builtFilePath)) {
    Write-Error "Built database path does not exist on disk: $builtFilePath"
    exit 1
}

if ($env:GITHUB_OUTPUT) {
    "builtPath=$builtFilePath" | Out-File -FilePath $env:GITHUB_OUTPUT -Append -Encoding utf8
}
