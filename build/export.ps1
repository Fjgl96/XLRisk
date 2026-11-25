<#
.SYNOPSIS
    Exports the VBA code of an XLRisk add-in to Source\.

.DESCRIPTION
    Use it after editing the code directly in the Excel VBA editor, so that the
    text files in Source\ (the files under version control) stay the single
    source of truth. Standard modules go to .bas, class modules to .cls and
    forms to .frm/.frx. The ThisWorkbook module is exported to
    ThisWorkbook.doccls for reference only: build.ps1 does not import it.

    Requirements: the same as build.ps1 (Windows, desktop Excel and
    "Trust access to the VBA project object model").

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File build\export.ps1 -Addin dist\XLRisk.xlam
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Addin
)

$ErrorActionPreference = 'Stop'
$root = Resolve-Path (Join-Path $PSScriptRoot '..')
$sourceDir = Join-Path $root 'Source'
$addinPath = Resolve-Path $Addin
# vbext_ComponentType -> file extension
$extensionByType = @{ 1 = '.bas'; 2 = '.cls'; 3 = '.frm'; 100 = '.doccls' }

$excel = $null
$wb = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $wb = $excel.Workbooks.Open($addinPath, 0, $true)   # read-only
    foreach ($component in $wb.VBProject.VBComponents) {
        $type = [int]$component.Type
        if (-not $extensionByType.ContainsKey($type)) { continue }
        if ($type -eq 100 -and $component.Name -ne 'ThisWorkbook') { continue }
        $file = Join-Path $sourceDir ($component.Name + $extensionByType[$type])
        $component.Export($file)
        Write-Host "  exported $($component.Name)"
    }
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
Write-Host "Review the changes with: git diff Source"
