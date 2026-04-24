# AFCA Piping BOM Generator

AFCA Piping BOM Generator is a C# SolidWorks BOM tool for piping assemblies. The solution separates **BomCore** (pure BOM logic) from **SolidWorksBOMAddin** (SolidWorks adapter/UI) and **BomPipeLauncher** (background SolidWorks automation for Explorer/PDM-style launch flows) so grouping, profile handling, diagnostics, and exports can be tested without a live SolidWorks session.

## Repository layout

- `src\BomCore` - SolidWorks-free BOM domain logic
- `src\SolidWorksBOMAddin` - COM add-in, adapters, and UI shell
- `src\BomPipeLauncher` - external launcher for background assembly processing
- `tests\BomCore.Tests` - unit tests for core logic
- `scripts\register-addin.ps1` / `scripts\unregister-addin.ps1` - COM host registration helpers
- `profiles\default.pipebom.json` - default AFCA pipe BOM profile
- `docs\` - architecture, rules, and test planning

## Current commands

Use the pinned SDK in `global.json` with the local dotnet executable:

```powershell
C:\Users\gombo\.dotnet\dotnet.exe restore .\SolidWorksBOMAddin.sln
C:\Users\gombo\.dotnet\dotnet.exe build .\SolidWorksBOMAddin.sln
C:\Users\gombo\.dotnet\dotnet.exe test .\SolidWorksBOMAddin.sln
C:\Users\gombo\.dotnet\dotnet.exe test .\tests\BomCore.Tests\BomCore.Tests.csproj --filter FullyQualifiedName~BomGeneratorTests
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

## Intended workflow

1. Open a SolidWorks assembly, or launch the tool from SolidWorks, Windows Explorer, or PDM.
2. In SolidWorks, use the **Pipe BOM > Open BOM Preview Shell** command to open the minimal WinForms preview/mapping shell.
3. Read selected-part properties or scan the full assembly through the SolidWorks adapter layer.
4. Load the effective JSON profile in this order: project-local, `%AppData%`, `%ProgramData%`, then built-in default.
5. Generate grouped BOM sections for pipe cuts, fittings, accessories, and other components.
6. Export the result to CSV or Excel.

The preview shell exposes buttons for selected-part property reads, active-assembly scans, BOM mapping edits, preview generation, and CSV/Excel export. `NumGaskets` and `NumClamps` are shown as accessory rules because they generate rows in **Pipe Accessories** rather than extra pipe columns.

The external launcher can be invoked directly against an assembly path:

```powershell
C:\Users\gombo\.dotnet\dotnet.exe run --project .\src\BomPipeLauncher\BomPipeLauncher.csproj -- --assembly "D:\AFCA\projects\test_project_2\Test_Project.SLDASM" --format csv
C:\Users\gombo\.dotnet\dotnet.exe run --project .\src\BomPipeLauncher\BomPipeLauncher.csproj -- --assembly "D:\AFCA\projects\test_project_2\Test_Project.SLDASM" --format csv --debug-report "D:\Exports\Test_Project.debug.json"
```

Pipe detection, grouping, ignored properties, and accessory generation are described in `docs\property-rules.md`.

## Installation and registration notes

- Current project targets: `.NET 8`, `net8.0`, and `net8.0-windows`
- Current SolidWorks interop packages are `32.1.0`, which correspond to the SolidWorks 2024 API generation
- The provided registration scripts write the add-in registration under `HKCU\Software\SolidWorks`, so the current-user registration flow does **not normally require admin rights**
- Installing SolidWorks itself, machine-wide prerequisites, or changing machine-wide registration policy may still require elevation on a workstation
- Effective profile lookup order is:
  1. profile beside the assembly
  2. `%AppData%\AFCA\SolidWorksBOMAddin\profiles\`
  3. `%ProgramData%\AFCA\SolidWorksBOMAddin\profiles\`
  4. built-in `profiles\default.pipebom.json`

## Validation note

`BomCore` can be validated with normal .NET tests, but SolidWorks-specific end-to-end validation depends on a local SolidWorks installation. The planned real-assembly validation target is `D:\AFCA\projects\test_project_2\Test_Project.SLDASM`.

The current launcher path has already been exercised against that fixture and produced:

- grouped CSV output with distinct pipe cut lengths
- a JSON debug report containing discovered properties and diagnostics
