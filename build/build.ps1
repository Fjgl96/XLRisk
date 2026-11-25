<#
.SYNOPSIS
    Builds dist\XLRisk.xlam from the VBA source files in Source\.

.DESCRIPTION
    1. Copies the base add-in (XLRisk.xlam at the repository root, which holds
       the ribbon, icons and the ThisWorkbook module) to dist\XLRisk.xlam.
    2. Opens the copy in a hidden Excel with macros disabled, removes every
       standard module, class module and form, saves and closes it.
    3. Reopens it, imports every .bas, .cls and .frm file from Source\ and saves.
    4. Checks that each source file became exactly one module and writes the
       SHA-256 checksum to dist\XLRisk.xlam.sha256.

    Requirements: Windows, desktop Excel, and the option
    "Trust access to the VBA project object model" enabled in
    File > Options > Trust Center > Trust Center Settings > Macro Settings.
    This script does not change that option; it only checks it.

    The built add-in is not digitally signed: building it removes the
    signature of the original XLRisk.xlam.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File build\build.ps1
#>
[CmdletBinding()]
param(
    [string]$BaseAddin,
    [string]$OutDir
)

$ErrorActionPreference = 'Stop'
$root = Resolve-Path (Join-Path $PSScriptRoot '..')
if (-not $BaseAddin) { $BaseAddin = Join-Path $root 'XLRisk.xlam' }
if (-not $OutDir) { $OutDir = Join-Path $root 'dist' }
$sourceDir = Join-Path $root 'Source'
$extensions = @('.bas', '.cls', '.frm')
# vbext_ComponentType: 1 = standard module, 2 = class module, 3 = form
$codeTypes = @(1, 2, 3)

$sourceFiles = @(Get-ChildItem -Path $sourceDir -File |
    Where-Object { $extensions -contains $_.Extension.ToLower() } |
    Sort-Object Name)
if ($sourceFiles.Count -eq 0) { throw "No .bas, .cls or .frm files found in $sourceDir" }

New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$target = Join-Path (Resolve-Path $OutDir) 'XLRisk.xlam'
Copy-Item -Path $BaseAddin -Destination $target -Force
Unblock-File -Path $target

Write-Host "Building $target from $($sourceFiles.Count) source files"

$excel = $null
$wb = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    # msoAutomationSecurityForceDisable: no macro of the add-in runs while it is built
    $excel.AutomationSecurity = 3

    # Pass 1: remove the old code modules
    $wb = $excel.Workbooks.Open($target)
    try {
        $components = $wb.VBProject.VBComponents
        $null = $components.Count
    }
    catch {
        throw ("Excel does not allow access to the VBA project. Enable " +
               "'Trust access to the VBA project object model' in File > Options > " +
               "Trust Center > Trust Center Settings > Macro Settings, then run the script again.")
    }
    $old = @($components | Where-Object { $codeTypes -contains [int]$_.Type })
    foreach ($component in $old) {
        Write-Host "  removing $($component.Name)"
        $components.Remove($component)
    }
    $wb.Save()
    $wb.Close($false)
    $wb = $null

    # Pass 2: import the source files
    $wb = $excel.Workbooks.Open($target)
    $components = $wb.VBProject.VBComponents
    foreach ($file in $sourceFiles) {
        Write-Host "  importing $($file.Name)"
        $null = $components.Import($file.FullName)
    }

    # Every source file must give exactly one module with the same name
    $names = @($components | Where-Object { $codeTypes -contains [int]$_.Type } |
        ForEach-Object { $_.Name })
    $expected = @($sourceFiles | ForEach-Object { $_.BaseName })
    $missing = @($expected | Where-Object { $names -notcontains $_ })
    $extra = @($names | Where-Object { $expected -notcontains $_ })
    if ($missing.Count -gt 0 -or $extra.Count -gt 0) {
        throw ("Module check failed. Missing: " + ($missing -join ', ') +
               ". Unexpected: " + ($extra -join ', '))
    }

    $wb.Save()
    $wb.Close($false)
    $wb = $null
}
finally {
    if ($wb -ne $null) { $wb.Close($false) }
    if ($excel -ne $null) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$hash = (Get-FileHash -Path $target -Algorithm SHA256).Hash.ToLower()
Set-Content -Path "$target.sha256" -Value "$hash  XLRisk.xlam" -Encoding ASCII
Write-Host "Done: $target"
Write-Host "SHA-256: $hash"
