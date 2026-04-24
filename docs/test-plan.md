# Test plan

## Goal

Validate the future AFCA Piping BOM Generator in layers so most behavior is proven without requiring a live SolidWorks session, while still reserving targeted workstation validation for the real adapter flow.

## Layer 1: `BomCore` unit tests

Primary automated coverage should live in `tests\BomCore.Tests` and run with `dotnet test`.

Planned unit-test focus areas:

- profile JSON deserialization from `profiles\default.pipebom.json`
- profile validation and duplicate-display-name diagnostics
- pipe detection based on `PipeLength`
- pipe grouping by `BOM` + `Pipe Identifier` + `Specification` + `PipeLength`
- separation of same-BOM pipes with different lengths
- ignored-property handling for gasket and ferrule helper fields
- accessory-row generation from `NumGaskets` and `NumClamps`
- export formatting and diagnostic propagation

These tests should not depend on SolidWorks interop or a local assembly file.

## Layer 2: profile and workflow validation

Repository-level validation should confirm that:

- the default profile shape matches the planning brief
- documentation and profile lookup order stay aligned
- invalid profile inputs lead to diagnostics and fallback behavior rather than unhandled exceptions

## Layer 3: SolidWorks adapter integration

Future integration validation should happen on a workstation with SolidWorks installed. This layer is expected to verify:

- reading properties from a selected part
- scanning an entire assembly recursively
- handling suppressed, lightweight, unresolved, virtual, and nested components without crashing
- producing grouped pipe and accessory rows that match the planning rules

## Real assembly validation target

The designated local validation assembly is:

`D:\AFCA\projects\test_project_2\Test_Project.SLDASM`

Planned checks against that assembly:

1. open the assembly in SolidWorks
2. read a known pipe part and confirm expected properties are visible
3. scan the full assembly and confirm pipe cuts stay separated by length
4. confirm accessory quantities are multiplied by grouped pipe quantity
5. export CSV/XLSX and inspect the resulting sections
6. invoke the launcher against the same assembly and confirm it can export without a manually opened SolidWorks session

## Environment note

SolidWorks-specific end-to-end validation is intentionally local-only. It should not be treated as a prerequisite for ordinary `BomCore` unit tests or for documentation/profile-only changes.
