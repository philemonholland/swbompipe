# Repository Guidelines

## Project Structure & Module Organization

This repository contains a C# SolidWorks BOM tool. Source projects live under `src/`:

- `src/BomCore`: SolidWorks-free BOM domain logic, exporters, profiles, diagnostics.
- `src/SolidWorksBOMAddin`: SolidWorks COM add-in, adapters, and WinForms shell.
- `src/BomPipeLauncher`: command-line/background export launcher.
- `src/BomPipePdmAddin` and `src/BomPipePdmVaultInstaller`: SolidWorks PDM integration.

Tests are in `tests/BomCore.Tests`. Profiles are in `profiles/`, and project documentation is in `docs/`. Installation and registration helpers are in `scripts/`. Avoid committing generated `bin/`, `obj/`, and `TestResults/` outputs.

## Build, Test, and Development Commands

Use the SDK pinned by `global.json`.

```powershell
dotnet restore .\SolidWorksBOMAddin.sln
dotnet build .\SolidWorksBOMAddin.sln
dotnet test .\SolidWorksBOMAddin.sln
dotnet test .\tests\BomCore.Tests\BomCore.Tests.csproj --filter FullyQualifiedName~BomGeneratorTests
```

`restore` downloads packages, `build` compiles all projects, and `test` runs the xUnit suite. Use `.\scripts\register-addin.ps1` and `.\scripts\unregister-addin.ps1` only after a successful build. Use `install-bompipe.cmd` and `uninstall-bompipe.cmd` for local install validation.

## Coding Style & Naming Conventions

Use C# with nullable reference types and implicit usings enabled. Follow existing style: four-space indentation, file-scoped namespaces, PascalCase for public types and members, camelCase for locals and parameters, and `Async` suffix for asynchronous methods. Keep `BomCore` independent from SolidWorks interop so unit tests remain machine-independent.

## Testing Guidelines

Tests use xUnit in `tests/BomCore.Tests`. Name test files after the target type, such as `BomGeneratorTests.cs`, and keep reusable fixtures in `TestData.cs`. Add or update unit tests for profile rules, grouping, exporters, diagnostics, and any behavior inside `BomCore`. SolidWorks and PDM flows require local manual validation; document those results in PR notes when touched.

## Commit & Pull Request Guidelines

Recent history uses short imperative summaries, for example `Enhance installation and uninstallation scripts...`. Keep commits focused and describe the user-visible change. Pull requests should include a concise summary, tests run, linked issue or task, and screenshots only when UI or installer prompts change. Call out SolidWorks, PDM vault, registry, or installer impacts explicitly.

## Security & Configuration Tips

Do not hard-code local vault names, assembly paths, credentials, or user directories. Keep profile defaults in `profiles/default.pipebom.json`; machine-specific overrides belong outside the repository.
