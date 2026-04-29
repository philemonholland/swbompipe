# Repository architecture

## Purpose

This repository is the scaffold for **AFCA Piping BOM Generator**, a SolidWorks BOM tool that reads piping assembly properties, applies a persistent JSON mapping profile, groups BOM rows, and exports human-facing CSV/XLSX output plus the BOMDB import JSON handoff.

## Top-level layout

- `SolidWorksBOMAddin.sln` - solution entry point
- `src\BomCore` - pure .NET class library for profiles, grouping, exporters, and diagnostics
- `src\SolidWorksBOMAddin` - SolidWorks COM add-in, adapters, and UI/orchestration
- `src\BomPipeLauncher` - external launcher for background BOM automation
- `src\BomPipePdmAddin` - PDM Professional add-in that contributes the vault context-menu command
- `src\BomPipePdmVaultInstaller` - vault registration utility for installing and removing the PDM add-in
- `tests\BomCore.Tests` - unit tests for `BomCore`
- `scripts\register-addin.ps1` / `scripts\unregister-addin.ps1` - SolidWorks add-in registration helpers
- `profiles\default.pipebom.json` - built-in default profile scaffold
- `docs\` - architecture, rules, and testing references

## Boundary rules

### `BomCore`

`BomCore` is the stable domain boundary. It owns:

- component/property records after extraction
- BOM profile models and validation
- grouping and accessory-row rules
- diagnostics and export-ready results

`BomCore` must stay free of SolidWorks interop references so it remains testable with ordinary .NET tests.

### `SolidWorksBOMAddin`

`SolidWorksBOMAddin` owns all SolidWorks-specific work:

- COM add-in registration and command surface
- reading assemblies, configurations, and component properties
- translating SolidWorks objects into `BomCore` records
- UI and launch/orchestration concerns

This boundary also covers non-embedded entry points. The product is expected to work not only from an in-session add-in command, but also from Windows Explorer and from a dedicated PDM Professional add-in/menu command.

### `BomPipeLauncher`

`BomPipeLauncher` is the background automation entry point for out-of-process launch flows. It is responsible for:

- starting SolidWorks through COM when launched from Explorer or the PDM add-in
- opening an assembly in the background
- reusing the SolidWorks adapter layer to read components
- loading the effective BOM profile and exporting CSV/XLSX output or the BOMDB import JSON file

For operator workflows, treat the `.bomdb.json` file as the authoritative downstream handoff to BOMDB. CSV and XLSX remain review-friendly exports for people, and BOMDB import may happen later on a different machine that does not run SolidWorks.

It should stay thin and delegate BOM rules to `BomCore` and SolidWorks object translation to the adapter project.

### `BomPipePdmAddin` and `BomPipePdmVaultInstaller`

`BomPipePdmAddin` is the vault-side integration boundary. It contributes the **Generate BOM with BOMPipe** command to the PDM Professional file context menu and launches the installed BOMPipe invoker for the selected assembly.

`BomPipePdmVaultInstaller` handles vault registration and removal for that add-in using the PDM add-in manager API. This keeps vault deployment concerns out of the pure launcher and SolidWorks add-in projects.

### `BomCore.Tests`

`BomCore.Tests` exists to validate core behavior without launching SolidWorks. Grouping, profile loading, diagnostics, and exporters should be proven here first.

## Data flow

1. The add-in or launcher starts from SolidWorks or an external launch surface.
2. The adapter layer reads an assembly or selected component from SolidWorks.
3. Adapter code resolves properties and produces core-friendly records.
4. `BomCore` classifies parts, applies the active profile, groups rows, and emits diagnostics.
5. Exporters or UI consume the resulting BOM sections.
6. Operators hand the `.bomdb.json` file to BOMDB, while CSV/XLSX stay available for human review.

## Profile boundary

Profiles are JSON and should be resolved in this order:

1. project-local profile beside the assembly
2. user profile in `%AppData%\AFCA\SolidWorksBOMAddin\profiles\`
3. company profile in `%ProgramData%\AFCA\SolidWorksBOMAddin\profiles\`
4. built-in default profile from this repository

Invalid or incomplete profile data should produce diagnostics and safe fallback behavior rather than crashing the run.

## Diagnostic posture

The repository direction is to prefer explicit diagnostics over hidden failures. Suppressed, unresolved, or otherwise unreadable components should be reported clearly while allowing the overall BOM generation flow to continue when possible.
