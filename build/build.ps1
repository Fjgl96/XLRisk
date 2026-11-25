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

Write-Host "build.ps1 revision 5 (2026-11-25)"
Write-Host "Building $target from $($sourceFiles.Count) source files"

# Excel only lets scripts change VBA code when "Trust access to the VBA project
# object model" is on (registry value AccessVBOM = 1). Check it before starting.
function Get-RegValue([string]$Path, [string]$Name) {
    try { return (Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop).$Name }
    catch { return $null }
}
$policyVbom = Get-RegValue 'HKCU:\Software\Policies\Microsoft\Office\16.0\Excel\Security' 'AccessVBOM'
$userVbom = Get-RegValue 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security' 'AccessVBOM'
if ($policyVbom -eq 0) {
    throw ("Your organization blocks 'Trust access to the VBA project object model' by policy. " +
           "Use the manual build in build\README.md. / Tu organizacion bloquea este acceso por politica: " +
           "usa el build manual de build\README.md.")
}
if ($policyVbom -ne 1 -and $userVbom -ne 1) {
    throw ("Enable 'Trust access to the VBA project object model' in Excel: File > Options > Trust Center > " +
           "Trust Center Settings > Macro Settings, then run the script again. / Activa 'Confiar en el acceso " +
           "al modelo de objetos de proyectos de VBA' en Excel: Archivo > Opciones > Centro de confianza > " +
           "Configuracion del Centro de confianza > Configuracion de macros, y vuelve a correr el script.")
}

# Returns the code modules (standard, class, form) of a VBA project as an array.
# The collection is walked with foreach: PowerShell cannot call VBComponents.Item
# with a number ("Index was outside the bounds of the array").
function Get-CodeComponents($Components) {
    $result = @()
    foreach ($component in $Components) {
        if ($codeTypes -contains [int]$component.Type) { $result += $component }
    }
    return ,$result
}

# Calls a COM member directly through IDispatch. PowerShell's own adapter turns
# the VBComponents collection into a fixed-size array, so Remove and Import fail
# ("The collection is of a fixed size") when called the normal way.
$getProperty = [System.Reflection.BindingFlags]::GetProperty
$invokeMethod = [System.Reflection.BindingFlags]::InvokeMethod
function Invoke-ComMember($Target, [string]$Name, [System.Reflection.BindingFlags]$Kind, [object[]]$Arguments) {
    [System.__ComObject].InvokeMember($Name, $Kind, $null, $Target, $Arguments)
}

# Returns the VBA project of a workbook. The function returns the project and not
# its VBComponents: PowerShell would unroll a returned collection into a plain
# fixed-size array, which has no working Remove or Import.
function Get-VBProject($Workbook) {
    $project = $Workbook.VBProject
    if ($null -eq $project) {
        throw ("Excel did not give access to the VBA project of $($Workbook.Name). Check that " +
               "'Trust access to the VBA project object model' is enabled and that the project is not locked.")
    }
    return $project
}

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
    $project = Get-VBProject $wb
    # Read the collection inline: returned from a function, PowerShell would unroll it
    $components = [System.__ComObject].InvokeMember('VBComponents', $getProperty, $null, $project, $null)
    if ($components -is [array]) { throw 'VBComponents arrived as a plain array; cannot edit the project.' }
    $old = Get-CodeComponents $components
    if ($old.Count -eq 0) { throw "No code modules found in $($BaseAddin): is it the XLRisk add-in?" }
    foreach ($component in $old) {
        Write-Host "  removing $($component.Name)"
        $null = Invoke-ComMember $components 'Remove' $invokeMethod @(,$component)
    }
    $wb.Save()
    $wb.Close($false)
    $wb = $null

    # Pass 2: import the source files
    $wb = $excel.Workbooks.Open($target)
    $project = Get-VBProject $wb
    # Read the collection inline: returned from a function, PowerShell would unroll it
    $components = [System.__ComObject].InvokeMember('VBComponents', $getProperty, $null, $project, $null)
    if ($components -is [array]) { throw 'VBComponents arrived as a plain array; cannot edit the project.' }
    foreach ($file in $sourceFiles) {
        Write-Host "  importing $($file.Name)"
        $null = Invoke-ComMember $components 'Import' $invokeMethod @($file.FullName)
    }

    # Every source file must give exactly one module with the same name
    $imported = Get-CodeComponents $components
    $names = @($imported | ForEach-Object { $_.Name })
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
