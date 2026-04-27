# Copilot Instructions

## Build, test, and lint

The repository is pinned to .NET SDK `8.0.420` via `global.json`. `BomCore` targets `net8.0`, the SolidWorks-facing projects target `net8.0-windows`, and the PDM Professional projects target `net48`.

```powershell
dotnet restore .\SolidWorksBOMAddin.sln
dotnet build .\SolidWorksBOMAddin.sln
dotnet test .\SolidWorksBOMAddin.sln
dotnet test .\tests\BomCore.Tests\BomCore.Tests.csproj --filter "FullyQualifiedName~BomCore.Tests.BomGeneratorTests.Generate_SeparatesPipesWithDifferentLengths"
```

There is no dedicated lint command checked into the repository today. Ordinary automated coverage lives in `tests\BomCore.Tests`; SolidWorks, launcher, and PDM flows still require local manual validation on a machine with SolidWorks and any needed vault setup.

## High-level architecture

`SolidWorksBOMAddin.sln` contains five product projects plus `tests\BomCore.Tests`:

- `src\BomCore` is the reusable domain layer. It owns profile serialization/loading, section classification, grouping, diagnostics, CSV/XLSX export, and debug-report generation.
- `src\SolidWorksBOMAddin` is the SolidWorks COM add-in and WinForms mapping shell. It exposes **Pipe BOM > Open BOM Preview Shell**, reads selected-part properties and active assemblies through SolidWorks interop, discovers properties, and lets users edit section columns, `Primary Family` section rules, accessory rules, and profile/settings files.
- `src\BomPipeLauncher` is the out-of-process automation entry point. It starts SolidWorks through COM, opens a `.SLDASM`, reuses the same reader + `BomCore` pipeline, and exports CSV/XLSX output plus an optional debug report.
- `src\BomPipePdmAddin` is the PDM Professional vault add-in. It contributes the **Generate BOM with BOMPipe** context-menu command for a selected assembly and launches the installed BOMPipe invoker.
- `src\BomPipePdmVaultInstaller` handles vault registration and removal for the PDM add-in across local vault views.
- `tests\BomCore.Tests` is the main automated safety net for profile serialization, section routing, grouping rules, profile fallback, exporters, and debug reports.

The shared flow is: read the assembly tree in SolidWorks, resolve effective custom properties into `ComponentRecord` instances, load the effective JSON profile, classify components into configured sections, run `BomGenerator`, then export BOM rows plus diagnostics/debug output. In-SolidWorks UI, launcher flows, and PDM flows are expected to reuse that pipeline instead of re-implementing BOM rules in multiple places.

The default profile is now a version 2 JSON model. `sectionColumnProfiles` and `sectionRules` drive per-section output for `Pipes`, `Tubes`, `Wires`, `Components`, `Connections`, `Fittings`, `Instruments`, `Systems`, and `Other`. The legacy `pipeColumns` array is still carried for pipe compatibility and migration, but section-based configuration is the current architecture.

## Key conventions

- Keep SolidWorks and PDM interop out of `BomCore`. `BomCore` must stay testable without SolidWorks and must not reference `SolidWorks.Interop.sldworks`, `SolidWorks.Interop.swconst`, or PDM interop packages.
- Effective profile lookup order is: `default.pipebom.json` beside the assembly, then `%AppData%\AFCA\SolidWorksBOMAddin\profiles\`, then `%ProgramData%\AFCA\SolidWorksBOMAddin\profiles\`, then the built-in `profiles\default.pipebom.json`. In the preview shell, a selected external settings file overrides that lookup and becomes the active save target. Invalid or unreadable profiles add diagnostics and fall back to the built-in default profile.
- SolidWorks property extraction precedence is configuration-specific first, then file-level, then component-level. The extractor keeps the first value found, so earlier scopes win.
- Assembly traversal must recurse through nested subassemblies. Suppressed components are skipped with info diagnostics; unresolved or unreadable models are skipped with warnings while traversal continues into any children. Virtual parts may not have stable file paths and fall back to component identity when grouped.
- Section routing is profile-driven. `Primary Family` is the default classifier for `Pipes`, `Tubes`, `Wires`, `Components`, `Connections`, `Fittings`, `Instruments`, `Systems`, and `Other`, and singular aliases such as `Fitting` normalize to the plural section. Missing or unmapped family values fall back to `Other`.
- Pipe-like sections depend on their length columns staying in the group key. `Pipes` use `PipeLength`, `Tubes` use `TubeLength`, and `Wires` use `WireLength`; the profile validator treats missing or non-grouped length columns as errors so equal BOM codes with different lengths do not merge.
- `NumGaskets` and `NumClamps` are accessory generators, not normal display columns. They produce separate accessory rows from grouped pipe quantities, and accessory sections normalize legacy `Pipe Accessories` values to `Other Accessories`.
- `BlueGasket`, `WhiteGasket`, `BlueFerrule`, and `WhiteFerrule` remain discoverable during property exploration and default-profile suggestion, but they are ignored by default as BOM columns.
- Prefer BOM rule changes in `BomCore` and prove them in `tests\BomCore.Tests` before touching the WinForms shell, SolidWorks add-in plumbing, launcher, or PDM integration.
- Host-side logging is best-effort. The add-in, PDM add-in, and vault installer write logs under `%LocalAppData%\AFCA\BOMPipe\logs`, and logging must never break the host process.
