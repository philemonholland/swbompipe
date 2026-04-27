# Copilot Instructions

## Build, test, and lint

The repository is pinned to .NET SDK `8.0.420` via `global.json`. `BomCore` targets `net8.0`, the SolidWorks-facing projects target `net8.0-windows`, and the PDM Professional projects target `net48`.

```powershell
dotnet restore .\SolidWorksBOMAddin.sln
dotnet build .\SolidWorksBOMAddin.sln
dotnet test .\SolidWorksBOMAddin.sln
dotnet test .\tests\BomCore.Tests\BomCore.Tests.csproj --filter "FullyQualifiedName~BomCore.Tests.BomGeneratorTests.Generate_SeparatesPipesWithDifferentLengths"
```

There is no dedicated lint command checked into the repository today. Ordinary automated coverage lives in `tests\BomCore.Tests`; SolidWorks and PDM flows still require local manual validation.

## High-level architecture

`SolidWorksBOMAddin.sln` contains five product projects plus `tests\BomCore.Tests`:

- `src\BomCore` is the reusable domain layer. It owns profile serialization/loading, property discovery, BOM grouping, diagnostics, CSV/XLSX export, and debug-report generation.
- `src\SolidWorksBOMAddin` is the SolidWorks COM add-in and WinForms shell. It exposes **Pipe BOM > Open BOM Preview Shell** and composes `SolidWorksAssemblyReader`, `SolidWorksSelectedComponentPropertyReader`, `ProfileStore`, `PropertyDiscoveryService`, and `BomGenerator`.
- `src\BomPipeLauncher` is the out-of-process automation entry point. It starts SolidWorks through COM, opens a `.SLDASM`, reuses the reader + `BomCore` pipeline, and exports CSV/XLSX output plus an optional debug report.
- `src\BomPipePdmAddin` is the PDM Professional vault add-in. It contributes the **Generate BOM with BOMPipe** context-menu command for a selected assembly and launches the installed BOMPipe invoker.
- `src\BomPipePdmVaultInstaller` handles vault registration and removal for the PDM add-in across local vault views.
- `tests\BomCore.Tests` is the main automated safety net for grouping rules, profile fallback, property discovery, exporters, and debug reports.

The shared flow is: read the assembly tree in SolidWorks, resolve effective custom properties into `ComponentRecord` instances, load the effective JSON profile, run `BomGenerator`, then export BOM rows plus diagnostics/debug output. In-SolidWorks UI, launcher flows, and PDM flows are expected to reuse that pipeline instead of re-implementing BOM rules in multiple places.

## Key conventions

- Keep SolidWorks and PDM interop out of `BomCore`. `BomCore` must stay testable without SolidWorks and must not reference `SolidWorks.Interop.sldworks`, `SolidWorks.Interop.swconst`, or PDM interop packages.
- Effective profile lookup order is: `default.pipebom.json` beside the assembly, then `%AppData%\AFCA\SolidWorksBOMAddin\profiles\`, then `%ProgramData%\AFCA\SolidWorksBOMAddin\profiles\`, then the built-in `profiles\default.pipebom.json`. Invalid or unreadable candidates add diagnostics and fall back to the built-in default profile.
- SolidWorks property extraction precedence is configuration-specific first, then file-level, then component-level. The extractor keeps the first value found, so earlier scopes win.
- Assembly traversal must recurse through nested subassemblies. Suppressed components are skipped with info diagnostics; unresolved or unreadable models are skipped with warnings while traversal continues into any children. Virtual parts may not have stable file paths and fall back to component identity when grouped.
- A component is treated as a pipe when `PipeLength` has a value. Default pipe grouping keeps rows distinct by `BOM`, `Pipe Identifier`, `Specification`, and `PipeLength`; same-BOM pipes with different lengths must not merge.
- `NumGaskets` and `NumClamps` are accessory generators, not normal display columns. They produce separate rows in **Pipe Accessories** using parsed numeric property value multiplied by component quantity across the grouped pipe components.
- `BlueGasket`, `WhiteGasket`, `BlueFerrule`, and `WhiteFerrule` remain discoverable during property exploration and default-profile suggestion, but they are ignored by default as BOM columns.
- Non-pipe grouping uses `BOM` when present; otherwise it falls back to file/configuration identity or virtual-component identity. Current `BomGenerator` output is pipe cuts, pipe accessories, and other components; do not assume a separate fittings classifier already exists just because `KnownBomSections` defines a `Fittings` section constant.
- Prefer adding or changing BOM rules in `BomCore` and proving them in `tests\BomCore.Tests` before touching the WinForms shell, SolidWorks add-in plumbing, launcher, or PDM integration.
