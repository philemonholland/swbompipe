# Copilot Instructions

This repository is still at the planning stage: the only tracked source file today is `README.md`, and the main implementation brief defines a future SolidWorks BOM add-in rather than an existing solution. Do not assume a buildable project already exists; scaffolding the solution may be part of the work.

## Build, test, and lint

No build, test, or lint commands are currently defined in the tracked repository. There is no `.sln`, project file, or existing test suite yet, so do not invent commands in follow-up work.

## High-level architecture

The planned product is a C# SolidWorks add-in named **AFCA Piping BOM Generator**. Its job is to scan a SolidWorks piping assembly, read component properties, apply a persistent property-mapping profile, generate grouped BOM rows, and export the result to CSV and XLSX.

Future work should not assume the only entry point is an in-session SolidWorks add-in command. The intended user experience also includes opening BOMPipe directly from Windows Explorer or from PDM by right-clicking an assembly, so keep launch/orchestration concerns separate from the core BOM pipeline.

Keep the architecture split once the solution is scaffolded:

- `SolidWorksBOMAddin`: COM add-in, UI shell, and all SolidWorks API integration
- `BomCore`: pure C# domain logic for profiles, grouping, diagnostics, and exporters
- `BomCore.Tests`: tests for `BomCore` without requiring SolidWorks

Keep SolidWorks API access behind adapters so the core BOM logic stays testable without launching SolidWorks. `BomCore` must not reference `SolidWorks.Interop.sldworks` or `SolidWorks.Interop.swconst`.

Persistent mapping profiles are JSON. The intended profile lookup order is:

1. Project-local profile beside the assembly
2. User profile in `%AppData%\AFCA\SolidWorksBOMAddin\profiles\`
3. Company profile in `%ProgramData%\AFCA\SolidWorksBOMAddin\profiles\`
4. Built-in default profile

The planned implementation order matters: scaffold first, then core models and profile loading, then grouping/property discovery/exporters, then SolidWorks adapters, then minimal UI, diagnostics, and installer work.

## Key conventions

- A component is considered a pipe when it has a `PipeLength` property.
- Pipe cut rows must be grouped by `BOM`, `Pipe Identifier`, `Specification`, and `PipeLength`; pipes with the same BOM but different lengths must stay separate.
- `NumGaskets` and `NumClamps` are accessory generators, not normal display columns. They produce separate accessory BOM rows using `property value * grouped pipe quantity`.
- Ignore these properties by default, but keep them discoverable in property exploration: `BlueGasket`, `WhiteGasket`, `BlueFerrule`, `WhiteFerrule`.
- Prefer explicit diagnostics over crashes. Invalid profiles and bad component reads should return diagnostics and continue with safe fallback behavior where the design brief calls for it.
- SolidWorks-specific reads should resolve properties in this order: configuration-specific, then file-level, then component-level if available.
- Suppressed components should be skipped by default and reported diagnostically rather than silently disappearing.
- Core logic should be implemented and tested before investing in UI polish; the riskiest work is assembly scanning, profile handling, grouping, and export behavior.
- Avoid coupling BOM generation to a manually opened SolidWorks session; design for both in-SolidWorks use and external launch flows such as Explorer or PDM context-menu entry points.
