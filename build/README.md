# Building XLRisk

The VBA code of XLRisk lives as text files in `Source/`. The add-in
`XLRisk.xlam` at the repository root is only the base that holds the ribbon,
the icons and the `ThisWorkbook` module; `build.ps1` replaces all of its code
with the files in `Source/`.

## Requirements

- Windows with desktop Excel (2013 or later).
- In Excel: File > Options > Trust Center > Trust Center Settings > Macro Settings >
  **Trust access to the VBA project object model**. The scripts only check this
  option; they never change it. If your organization does not allow it, see
  "Manual build" below.

## Build

```powershell
powershell -ExecutionPolicy Bypass -File build\build.ps1
```

The script writes `dist\XLRisk.xlam` and `dist\XLRisk.xlam.sha256`, and stops with
an error if any source file did not become a module.

Then load `dist\XLRisk.xlam` in Excel (Developer > Excel Add-ins > Browse) and run
the self tests (`tests/README.md`).

## Edit in the VBA editor, then export

If you change the code in the Excel VBA editor, export it back to `Source/` before
committing:

```powershell
powershell -ExecutionPolicy Bypass -File build\export.ps1 -Addin dist\XLRisk.xlam
git diff Source
```

## Manual build (without access to the VBA object model)

1. Copy `XLRisk.xlam` to a new folder and open it in Excel.
2. Open the VBA editor (Alt+F11). In the XLRisk project remove every module, class
   module and form (right click > Remove, answer No to export).
3. File > Import File, and import every `.bas`, `.cls` and `.frm` file of `Source/`.
4. Save, close Excel and load the new add-in.

## Digital signature

The original `XLRisk.xlam` is signed by its author. A built add-in is not signed.
For personal use, sign it with a self-signed certificate (Office tool SelfCert) or
load it from a trusted location. To distribute it, use a code-signing certificate.
