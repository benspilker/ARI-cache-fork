<#
.SYNOPSIS
    Replaces the buggy upstream ImportExcel Set-ExcelRange.ps1 with the vendored
    patched version from this ARI-cache-fork repo.

.DESCRIPTION
    ImportExcel 7.8.10 + EPPlus 6/7 on Linux/.NET 9 throws "The property 'X' cannot
    be found on this object" errors when PowerShell tries to set Style properties
    (HorizontalAlignment, VerticalAlignment, WrapText, Numberformat.Format, Font.*,
    Border.*, Fill.*, Locked, etc.). This is a known issue tracked at
    https://github.com/Azure/ARI/issues/17 and is unfixed in ImportExcel as of 2026-08.

    The fix in this repo wraps every $Range.Style.* assignment in try/catch so the
    Excel generation completes (with possibly less styling on Linux) instead of crashing.

    This installer is idempotent:
      - Backs up the original once (Set-ExcelRange.ps1.bak)
      - Marker-checks the patched file before replacing
      - Skips silently if already applied
      - Walks all installed ImportExcel versions under the cache root

.PARAMETER CacheRoot
    Override the cache root. Defaults to '/tmp/windmill/cache/powershell/ImportExcel'
    which is the path Windmill's PowerShell worker uses on Linux.

.PARAMETER Force
    Re-apply even if the marker is found (e.g. to revert to a fresh patched copy).

.EXAMPLE
    # Default invocation (recommended):
    ./Apply-VendoredImportExcelPatch.ps1

.EXAMPLE
    # Force re-apply (e.g. after updating this repo's Public/Set-ExcelRange.ps1):
    ./Apply-VendoredImportExcelPatch.ps1 -Force

.NOTES
    File:        Apply-VendoredImportExcelPatch.ps1
    Lives at:    ARI-cache-fork/Modules/Vendored/ImportExcel/Apply-VendoredImportExcelPatch.ps1
    Marker:      # HA-LINUX-PATCH-v1
    Author:      Benjamin Spilker / HA assistant
    Created:     2026-08-17
    Refs:        https://github.com/Azure/ARI/issues/17
                 dfinke/ImportExcel@master/Public/Set-ExcelRange.ps1
#>
[CmdletBinding()]
param(
    [string]$CacheRoot = '/tmp/windmill/cache/powershell/ImportExcel',
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$marker = 'HA-LINUX-PATCH-v1'

# Patched source: this script lives at <RepoRoot>/Modules/Vendored/ImportExcel/
# Patched file:  <RepoRoot>/Modules/Vendored/ImportExcel/Public/Set-ExcelRange.ps1
$patchedSrc = Join-Path $PSScriptRoot 'Public/Set-ExcelRange.ps1'

function Write-Step { param([string]$Msg) Write-Host "[VendoredPatch] $Msg" -ForegroundColor Cyan }
function Write-OK   { param([string]$Msg) Write-Host "[VendoredPatch] $Msg" -ForegroundColor Green }
function Write-Skip { param([string]$Msg) Write-Host "[VendoredPatch] $Msg" -ForegroundColor DarkGray }
function Write-Warn { param([string]$Msg) Write-Warning "[VendoredPatch] $Msg" }

# 1. Verify patched source exists in the fork
if (-not (Test-Path -LiteralPath $patchedSrc)) {
    throw "Patched source not found at: $patchedSrc`nMake sure this script is co-located with Public/Set-ExcelRange.ps1 in your ARI-cache-fork repo."
}

# 2. Verify patched source contains the marker (catches accidentally-wrong files)
$patchedContent = Get-Content -LiteralPath $patchedSrc -Raw
if ($patchedContent -notmatch [regex]::Escape($marker)) {
    throw "Patched source at $patchedSrc does not contain marker '$marker'. Refusing to install a file that may not be the right one."
}

# 3. Find ImportExcel cache root
if (-not (Test-Path -LiteralPath $CacheRoot)) {
    Write-Skip "ImportExcel cache root not found at $CacheRoot (ImportExcel not yet installed?). Nothing to patch; will run on next Windmill run that triggers Install-Module ImportExcel."
    return
}

# 4. Iterate all installed ImportExcel versions (7.8.10, 7.8.11, etc.)
$installedVersions = Get-ChildItem -LiteralPath $CacheRoot -Directory -ErrorAction SilentlyContinue
if (-not $installedVersions) {
    Write-Skip "No installed ImportExcel versions found under $CacheRoot"
    return
}

foreach ($vdir in $installedVersions) {
    $target = Join-Path $vdir.FullName 'Public/Set-ExcelRange.ps1'

    if (-not (Test-Path -LiteralPath $target)) {
        Write-Skip "$($vdir.Name): Set-ExcelRange.ps1 not present at $target (skipping)"
        continue
    }

    $targetContent = Get-Content -LiteralPath $target -Raw

    # Already patched?
    if (-not $Force -and ($targetContent -match [regex]::Escape($marker))) {
        Write-Skip "$($vdir.Name): already patched (marker found)"
        continue
    }

    # Back up the original once
    $backup = "$target.bak"
    if (-not (Test-Path -LiteralPath $backup)) {
        Copy-Item -LiteralPath $target -Destination $backup -Force
        Write-Step "$($vdir.Name): backed up original to $backup"
    } else {
        Write-Step "$($vdir.Name): backup already exists at $backup"
    }

    # Replace with vendored patched version
    Copy-Item -LiteralPath $patchedSrc -Destination $target -Force
    Write-OK "$($vdir.Name): replaced Set-ExcelRange.ps1 with vendored patched version"
}

Write-OK "Done. Restart any PowerShell sessions so the patched module is re-imported."
