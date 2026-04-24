# Repository architecture

## Purpose

This repository is the scaffold for **AFCA Piping BOM Generator**, a SolidWorks BOM tool that reads piping assembly properties, applies a persistent JSON mapping profile, groups BOM rows, and exports CSV/XLSX output.

## Top-level layout

- `SolidWorksBOMAddin.sln` - solution entry point
- `src\BomCore` - pure .NET class library for profiles, grouping, exporters, and diagnostics
- `src\SolidWorksBOMAddin` - SolidWorks COM add-in, adapters, and UI/orchestration
- `src\BomPipeLauncher` - external launcher for Explorer/PDM-style background automation
- `tests\BomCore.Tests` - unit tests for `BomCore`
- `scripts\register-addin.ps1` / `scripts\unregister-addin.ps1` - COM registration helpers
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

This boundary also covers non-embedded entry points. The product is expected to work not only from an in-session add-in command, but also from Windows Explorer or PDM launch flows.

### `BomPipeLauncher`

`BomPipeLauncher` is the background automation entry point for out-of-process launch flows. It is responsible for:

- starting SolidWorks through COM when launched from Explorer or PDM
- opening an assembly in the background
- reusing the SolidWorks adapter layer to read components
- loading the effective BOM profile and exporting CSV/XLSX output

It should stay thin and delegate BOM rules to `BomCore` and SolidWorks object translation to the adapter project.

### `BomCore.Tests`

`BomCore.Tests` exists to validate core behavior without launching SolidWorks. Grouping, profile loading, diagnostics, and exporters should be proven here first.

## Data flow

1. The add-in or launcher starts from SolidWorks or an external launch surface.
2. The adapter layer reads an assembly or selected component from SolidWorks.
3. Adapter code resolves properties and produces core-friendly records.
4. `BomCore` classifies parts, applies the active profile, groups rows, and emits diagnostics.
5. Exporters or UI consume the resulting BOM sections.

## Profile boundary

Profiles are JSON and should be resolved in this order:

1. project-local profile beside the assembly
2. user profile in `%AppData%\AFCA\SolidWorksBOMAddin\profiles\`
3. company profile in `%ProgramData%\AFCA\SolidWorksBOMAddin\profiles\`
4. built-in default profile from this repository

Invalid or incomplete profile data should produce diagnostics and safe fallback behavior rather than crashing the run.

## Diagnostic posture

The repository direction is to prefer explicit diagnostics over hidden failures. Suppressed, unresolved, or otherwise unreadable components should be reported clearly while allowing the overall BOM generation flow to continue when possible.
