# AFCA Piping BOM Generator

AFCA Piping BOM Generator is a C# SolidWorks BOM tool for piping assemblies. The solution separates **BomCore** (pure BOM logic), **SolidWorksBOMAddin** (SolidWorks adapter/UI), **BomPipeLauncher** (background export launcher), and the new **BomPipePdmAddin** / **BomPipePdmVaultInstaller** pieces for SolidWorks PDM Professional integration so grouping, profile handling, diagnostics, and exports can be tested without a live SolidWorks session.

## Repository layout

- `src\BomCore` - SolidWorks-free BOM domain logic
- `src\SolidWorksBOMAddin` - COM add-in, adapters, and UI shell
- `src\BomPipeLauncher` - external launcher for background assembly processing
- `src\BomPipePdmAddin` - SolidWorks PDM Professional add-in that contributes the right-click menu command
- `src\BomPipePdmVaultInstaller` - vault registration utility for installing or removing the PDM add-in
- `tests\BomCore.Tests` - unit tests for core BOM rules and exporters
- `tests\BomPipeLauncher.Tests` - unit tests for launcher argument parsing and export-path behavior
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

## Export outputs

- `.bomdb.json` is the authoritative BOMDB handoff artifact. Import this JSON file into BOMDB. BOMPipe shows editable `Project` and `Project Name` fields at the top of the preview shell, fills them from the root assembly custom properties when available, and writes them to the JSON root as `project` and `project_name`, plus `assembly_path` for traceability.
- `.csv` and `.xlsx` are human-facing BOM outputs for review, sharing, and shop-floor use. They are not the BOMDB import contract.
- `.debug.json` is a troubleshooting artifact for diagnostics and discovered-property investigation.
- BOMDB can run on a different machine from the SolidWorks/BOMPipe export machine. SolidWorks is only needed where the assembly is read and exported.

## Operator workflow

1. Start from one of the supported export surfaces:
   - **In SolidWorks UI:** open the assembly, then use **Pipe BOM > Open BOM Preview Shell**.
   - **Launcher / automation:** run `BomPipeLauncher` against a `.SLDASM`.
   - **Explorer / PDM:** use **Generate BOM with BOMPipe** on an installed workstation.
2. Let BOMPipe read the assembly, resolve the active profile, and generate the grouped BOM.
3. Export the output you need:
   - use **Export BOMDB JSON** in the preview shell when the next step is BOMDB import
   - use **Export CSV** or **Export Excel** when the next step is human review or distribution
   - use launcher sidecar output when you want both a human-facing workbook and the BOMDB import file from one run
4. Hand the resulting `.bomdb.json` file to the BOMDB import workflow. This import can happen on another machine that does not have SolidWorks installed.
5. Keep the `.xlsx` / `.csv` file for operators and the `.debug.json` file for troubleshooting if questions come up during import or review.

The preview shell exposes buttons for selected-part property reads, active-assembly scans, BOM mapping edits, preview generation, **Export CSV**, **Export Excel**, and **Export BOMDB JSON**. Those preview/export actions are available both in the top command strip and directly inside the **BOM Mapping** workspace so the BOMDB handoff path stays visible while editing mappings. The Project field is required for BOMDB JSON export so BOMDB can create or update the correct project automatically. The BOMDB export dialog suggests a `.bomdb.json` file name. `NumGaskets` and `NumClamps` are shown as accessory rules because they generate rows in **Pipe Accessories** rather than extra pipe columns.

The installed manager shows a build line such as `Build: <revision> | Release | Installed: <timestamp>` in the header. If that line says the manifest is missing or invalid, reinstall BOMPipe before checking UI/export behavior; the installer writes this manifest after publishing the current source.

The external launcher can be invoked directly against an assembly path:

```powershell
dotnet run --project .\src\BomPipeLauncher\BomPipeLauncher.csproj -- --assembly "C:\Path\To\Assembly.SLDASM" --format csv
dotnet run --project .\src\BomPipeLauncher\BomPipeLauncher.csproj -- --assembly "C:\Path\To\Assembly.SLDASM" --format xlsx --output "C:\Path\To\Exports\Assembly.xlsx" --bomdb-output "C:\Path\To\Exports\Assembly.bomdb.json"
dotnet run --project .\src\BomPipeLauncher\BomPipeLauncher.csproj -- --assembly "C:\Path\To\Assembly.SLDASM" --format bomdb-json --output "C:\Path\To\Exports\Assembly.bomdb.json"
dotnet run --project .\src\BomPipeLauncher\BomPipeLauncher.csproj -- --assembly "C:\Path\To\Assembly.SLDASM" --format csv --debug-report "C:\Path\To\Exports\Assembly.debug.json"
```

Typical launcher choices:

- `--format bomdb-json --output ...bomdb.json` when the run is only for BOMDB import.
- `--format xlsx --output ...xlsx --bomdb-output ...bomdb.json` when the run should produce both the operator workbook and the BOMDB import file.
- `--debug-report ...debug.json` when diagnostics should be saved alongside the export.

When `--bomdb-output` is used, `BomPipeLauncher` writes both the primary export path and the BOMDB JSON path to the console on success.

Installed Explorer and PDM flows use the second pattern automatically: they export an `.xlsx` workbook for operators, a `.bomdb.json` file for BOMDB import, and a `.debug.json` report for troubleshooting. The installed invoker shows all three output paths after export.

Pipe detection, grouping, ignored properties, and accessory generation are described in `docs\property-rules.md`.

## BOMDB JSON export contract

`BomCore` now includes a dedicated BOMDB-oriented payload model (`BomDbImportFile`) plus reusable creation/serialization helpers (`BomDbExportService` and `BomDbJsonExporter`). The authoritative handoff is a JSON object rooted at `contract_version`, optional project metadata (`project`, `project_name`, `assembly_path`), and `rows`. Each row uses BOMDB's canonical import fields (`file_path`, `configuration_name`, `quantity`, `component_name`, `part_number`, `description`, `material`, `item_number`, and `custom_properties_json`) instead of the older structured debug-style schema. CSV/XLSX remain separate human-readable exports.

## Installation and registration notes

- Current project targets: `.NET 8`, `net8.0`, `net8.0-windows`, and `net48` for the PDM Professional add-in components
- Current SolidWorks interop packages are `32.1.0`, which correspond to the SolidWorks 2024 API generation
- The SolidWorks add-in registration uses per-user COM registration under `HKCU\Software\Classes` plus a per-user background loader that calls `LoadAddIn("AFCA.PipingBom.Generator")` for running SolidWorks sessions, so the install can stay no-admin
- `install-bompipe.cmd` publishes the current source into a per-user install under `%LocalAppData%\AFCA\BOMPipe`, registers the SolidWorks add-in, starts the per-user SolidWorks loader, installs `Generate BOM with BOMPipe` for `.SLDASM` in Explorer, and registers the PDM Professional add-in in the selected vaults
- The installed right-click command runs `Invoke-BOMPipe.ps1`, which exports an `.xlsx` BOM, a `.bomdb.json` BOMDB import file, and a `.debug.json` report to `%UserProfile%\Documents\AFCA\BOMPipe\Exports`
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
- a `.bomdb.json` BOMDB import file
- a JSON debug report containing discovered properties and diagnostics
