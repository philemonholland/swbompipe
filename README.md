# AFCA Piping BOM Generator

AFCA Piping BOM Generator is a C# SolidWorks BOM tool for piping assemblies. The solution separates **BomCore** (pure BOM logic), **SolidWorksBOMAddin** (SolidWorks adapter/UI), **BomPipeLauncher** (background export launcher), and the new **BomPipePdmAddin** / **BomPipePdmVaultInstaller** pieces for SolidWorks PDM Professional integration so grouping, profile handling, diagnostics, and exports can be tested without a live SolidWorks session.

## Repository layout

- `src\BomCore` - SolidWorks-free BOM domain logic
- `src\SolidWorksBOMAddin` - COM add-in, adapters, and UI shell
- `src\BomPipeLauncher` - external launcher for background assembly processing
- `src\BomPipePdmAddin` - SolidWorks PDM Professional add-in that contributes the right-click menu command
- `src\BomPipePdmVaultInstaller` - vault registration utility for installing or removing the PDM add-in
- `tests\BomCore.Tests` - unit tests for core logic
- `scripts\register-addin.ps1` / `scripts\unregister-addin.ps1` - SolidWorks add-in registration helpers
- `profiles\default.pipebom.json` - default AFCA pipe BOM profile
- `docs\` - architecture, rules, and test planning

## Current commands

Use the pinned SDK in `global.json` with a .NET 8 SDK on your machine:

```powershell
dotnet restore .\SolidWorksBOMAddin.sln
dotnet build .\SolidWorksBOMAddin.sln
dotnet test .\SolidWorksBOMAddin.sln
dotnet test .\tests\BomCore.Tests\BomCore.Tests.csproj --filter FullyQualifiedName~BomGeneratorTests
```

Register or unregister the add-in host after building:

```powershell
.\scripts\register-addin.ps1
.\scripts\unregister-addin.ps1
```

To register the add-in at SolidWorks startup for the current user:

```powershell
.\scripts\register-addin.ps1 -StartAtSolidWorksStartup
```

One-click install and uninstall from the repository root:

```cmd
install-bompipe.cmd
uninstall-bompipe.cmd
```

## Intended workflow

1. Open a SolidWorks assembly, or launch the tool from SolidWorks, Windows Explorer, or the PDM Professional vault context menu.
2. In SolidWorks, use the **Pipe BOM > Open BOM Preview Shell** command to open the minimal WinForms preview/mapping shell.
3. Read selected-part properties or scan the full assembly through the SolidWorks adapter layer.
4. Load the effective JSON profile in this order: project-local, `%AppData%`, `%ProgramData%`, then built-in default.
5. Generate grouped BOM sections for pipe cuts, fittings, accessories, and other components.
6. Export the result to CSV or Excel.

The preview shell exposes buttons for selected-part property reads, active-assembly scans, BOM mapping edits, preview generation, and CSV/Excel export. `NumGaskets` and `NumClamps` are shown as accessory rules because they generate rows in **Pipe Accessories** rather than extra pipe columns.

The external launcher can be invoked directly against an assembly path:

```powershell
dotnet run --project .\src\BomPipeLauncher\BomPipeLauncher.csproj -- --assembly "C:\Path\To\Assembly.SLDASM" --format csv
dotnet run --project .\src\BomPipeLauncher\BomPipeLauncher.csproj -- --assembly "C:\Path\To\Assembly.SLDASM" --format csv --debug-report "C:\Path\To\Exports\Assembly.debug.json"
```

Pipe detection, grouping, ignored properties, and accessory generation are described in `docs\property-rules.md`.

## Installation and registration notes

- Current project targets: `.NET 8`, `net8.0`, `net8.0-windows`, and `net48` for the PDM Professional add-in components
- Current SolidWorks interop packages are `32.1.0`, which correspond to the SolidWorks 2024 API generation
- The SolidWorks add-in registration uses per-user COM registration under `HKCU\Software\Classes` plus a per-user background loader that calls `LoadAddIn("AFCA.PipingBom.Generator")` for running SolidWorks sessions, so the install can stay no-admin
- `install-bompipe.cmd` stages a per-user install under `%LocalAppData%\AFCA\BOMPipe`, registers the SolidWorks add-in, starts the per-user SolidWorks loader, installs `Generate BOM with BOMPipe` for `.SLDASM` in Explorer, and registers the PDM Professional add-in in the selected vaults
- The installed right-click command runs `Invoke-BOMPipe.ps1`, which exports an `.xlsx` BOM plus a `.debug.json` report to `%UserProfile%\Documents\AFCA\BOMPipe\Exports`
- PDM Professional integration is now a real `IEdmAddIn5` vault add-in, not just an Explorer shell verb; vault registration requires a PDM admin-capable login for the target vault view
- Use `-PdmVaultName <vault>` to target a specific local vault view during install or uninstall; if omitted, the vault installer targets all locally registered vault views
- Effective profile lookup order is:
  1. profile beside the assembly
  2. `%AppData%\AFCA\SolidWorksBOMAddin\profiles\`
  3. `%ProgramData%\AFCA\SolidWorksBOMAddin\profiles\`
  4. built-in `profiles\default.pipebom.json`

Installer examples:

```powershell
.\scripts\install-bompipe.ps1
.\scripts\install-bompipe.ps1 -PdmVaultName AFCA
.\scripts\install-bompipe.ps1 -ForceRebuild
.\scripts\install-bompipe.ps1 -SkipPdmRegistration
.\scripts\uninstall-bompipe.ps1
.\scripts\uninstall-bompipe.ps1 -PdmVaultName AFCA
.\scripts\register-addin.ps1 -StartAtSolidWorksStartup
```

## Validation note

`BomCore` can be validated with normal .NET tests, but SolidWorks-specific end-to-end validation depends on a local SolidWorks installation, a local `.SLDASM` fixture that exists on the target machine, and a local PDM vault view when validating the PDM Professional menu.

The current launcher path has already been exercised against that fixture and produced:

- grouped CSV output with distinct pipe cut lengths
- a JSON debug report containing discovered properties and diagnostics
