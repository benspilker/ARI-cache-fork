# Vendored ImportExcel Patch (Linux EPPlus workaround)

This directory contains a vendored, patched version of **one file** from the
[ImportExcel](https://github.com/dfinke/ImportExcel) PowerShell module
(`Public/Set-ExcelRange.ps1`), plus the installer that applies it.

## Why this exists

`ImportExcel 7.8.10` (the version Windmill installs by default) calls EPPlus
property setters like:

```powershell
$Range.Style.HorizontalAlignment = $HorizontalAlignment
$Range.Style.Numberformat.Format = 'm/d/yy h:mm'
$Range.Style.Font.Bold           = $true
```

On **Windows + .NET Framework**, these work fine. On **Linux + .NET 9** (the
runtime used by the Windmill PowerShell worker), EPPlus 6/7's `ExcelStyle`
object does **not** expose those setters the same way, and PowerShell's
reflection layer throws:

> The property 'X' cannot be found on this object.

The first crash site is `Set-ExcelRange.ps1:107`
(`$Range.Style.HorizontalAlignment = $HorizontalAlignment`), called from
`Export-Excel` whenever a sheet is auto-styled.

This is a known issue:
- [Azure/ARI issue #17](https://github.com/Azure/ARI/issues/17) — open since 2021
- ImportExcel itself has not released a fix as of 2026-08
- It only affects Linux workers; the ARI-cache-fork Excel generation works
  fine on Windows

Rather than ship a runtime monkey-patch (hacky, fragile, per-run overhead),
we vendor the one file that needs changing and apply it **once at install
time** as part of the ARI setup flow.

## What's in this directory

```
Vendored/ImportExcel/
├── README.md                              ← you are here
├── Apply-VendoredImportExcelPatch.ps1     ← idempotent installer (call this once)
└── Public/
    └── Set-ExcelRange.ps1                 ← vendored patched file (try/catch wrappers)
```

The patched `Set-ExcelRange.ps1` is identical to upstream
[dfinke/ImportExcel@master/Public/Set-ExcelRange.ps1](https://github.com/dfinke/ImportExcel/blob/master/Public/Set-ExcelRange.ps1)
**except** that every `$Range.Style.* = ...` assignment is wrapped in
`try { ... } catch { Write-Verbose "[HA-LinuxPatch] ..." }`. The patch
marker comment `# HA-LINUX-PATCH-v1` is included on each wrapper.

## How it's invoked from Windmill

The Windmill script `f/ARI/csp_or_operate_ari_batched_script` (or the
equivalent `csp_or_operate_ari_batched_script.ps1` in `tools-for-ari-git`)
calls this installer once after `Install-Module ImportExcel`:

```powershell
Install-Module ImportExcel -RequiredVersion 7.8.10 -Force -SkipPublisherCheck

# Apply the vendored Linux patch on top of the installed module
& "<ariRepo>/Modules/Vendored/ImportExcel/Apply-VendoredImportExcelPatch.ps1"
```

The installer:

1. Verifies the patched source exists and contains the `# HA-LINUX-PATCH-v1`
   marker (catches accidentally-wrong files)
2. Walks `/tmp/windmill/cache/powershell/ImportExcel/<version>/` for every
   installed ImportExcel version
3. Backs up the original once (`Set-ExcelRange.ps1.bak`)
4. Replaces the upstream file with the vendored patched version
5. Skips silently if already patched (marker check)

The next PowerShell session that imports ImportExcel will pick up the
patched version automatically (PowerShell caches modules per-session, so
in-flight sessions need to be restarted — the installer prints a reminder).

## What happens on Windows

The patch is **safe to apply on Windows** too — the `try { ... } catch { }`
wrappers just become no-ops because the assignments succeed. There's no
Windows-specific check; we always apply the patch. This keeps the
Windmill script body identical across both operating systems.

## Trade-offs accepted

| Trade-off | Detail |
|---|---|
| **Styling loss on Linux** | When EPPlus rejects a Style setter, we silently skip it (logged via `Write-Verbose`). Excel sheets will be less visually polished than on Windows, but they will **complete generation instead of crashing**. |
| **Manual re-vendoring when ImportExcel releases a fix** | When ImportExcel ships a Linux fix in a future version, we'll need to: (1) re-download the upstream `Set-ExcelRange.ps1`, (2) verify no `.Style.*=` assignments remain, (3) delete this entire `Vendored/ImportExcel/` directory, (4) bump `-RequiredVersion` in the Windmill script. |
| **We don't get other ImportExcel bug fixes automatically** | Until this directory is deleted, we're locked to the version of `Set-ExcelRange.ps1` we vendored. New ImportExcel releases won't reach us until we re-vendor. |

## When to delete this directory

Delete `Vendored/ImportExcel/` (and remove the installer call from the
Windmill script body) when **any** of these become true:

1. ImportExcel releases a version where the Linux Style-setter bug is fixed
   (track via [Azure/ARI issue #17](https://github.com/Azure/ARI/issues/17))
2. We migrate the Windmill PowerShell worker off Linux/.NET 9 (back to
   Windows/.NET Framework, or to a newer PowerShell runtime where EPPlus
   handles Style setters correctly)
3. We replace ImportExcel with a different Excel-generation library entirely

## Updating the patched file (advanced)

If you need to update the vendored `Set-ExcelRange.ps1` to a newer upstream
version:

```bash
# 1. Download the new upstream version
curl -sSL -o Public/Set-ExcelRange.ps1 \
  https://raw.githubusercontent.com/dfinke/ImportExcel/master/Public/Set-ExcelRange.ps1

# 2. Re-apply the try/catch wrappers
#    (Use the same regex: $Range.Style.<id>(.<id>)*<whitespace>=<not-equal>)
#    Each wrapped assignment gets a "# HA-LINUX-PATCH-v1" marker comment.

# 3. Bump the marker version in both the file AND Apply-VendoredImportExcelPatch.ps1
#    (so the installer can detect old patches and re-apply)

# 4. Commit and push. Next Windmill run will use -Force to upgrade.
```

## File provenance

- **Source repo**: https://github.com/dfinke/ImportExcel
- **Source file**: `Public/Set-ExcelRange.ps1`
- **License**: Apache License 2.0 (ImportExcel) — patch is a derivative work
  under the same license
- **Patch author**: Benjamin Spilker / HA assistant, 2026-08-17
- **Refs**: https://github.com/Azure/ARI/issues/17
