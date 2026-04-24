# Agent build plan: SolidWorks Piping BOM Add-in

## Goal

Build a SolidWorks C# add-in that scans a piping assembly, reads component properties, applies a persistent property-mapping profile, and generates a BOM containing:

- pipe cut lengths
- fitting quantities
- gasket quantities
- clamp quantities
- all selected user-defined part properties
- grouped BOM rows
- export to CSV and Excel
- later: optional insertion into a SolidWorks drawing/table

The add-in must let the user select a part, read its custom properties, choose which properties to include, rename them as BOM table headers, and save that configuration permanently.

---

## Main architecture rule

Do not let the SolidWorks API contaminate the whole codebase.

Keep SolidWorks API access behind adapters. The BOM logic must be testable without opening SolidWorks.

Architecture:

```text
SolidWorksBOMAddin/
  src/
    SolidWorksBOMAddin/          # COM add-in, UI, SolidWorks API adapter
    BomCore/                     # pure C# BOM logic, no SolidWorks dependency
    BomCore.Tests/               # unit tests for grouping, mapping, profile loading
  profiles/
    default.pipebom.json
  docs/
    architecture.md
    property-rules.md
    test-plan.md
```

The core library must not reference:

```text
SolidWorks.Interop.sldworks
SolidWorks.Interop.swconst
```

Only the add-in/adapter project should reference those.

---

## BOM rules to implement

### Pipe detection

A component is a pipe if it has this property:

```text
PipeLength
```

### Pipe properties currently known

```text
Pipe Identifier
SWbompartno
Specification
BlueGasket
WhiteGasket
NumGaskets
PipeLength
NumClamps
Date
Author
WhiteFerrule
BlueFerrule
BOM
```

### Ignored properties

Ignore these by default:

```text
BlueGasket
WhiteGasket
WhiteFerrule
BlueFerrule
```

They can still appear in the property explorer, but they should not be selected by default.

### Pipe cut-list grouping

Pipe cut rows must be grouped by:

```text
BOM
Pipe Identifier
Specification
PipeLength
```

Do not group pipes only by `BOM`, because two pipes with the same BOM code but different lengths are different cut pieces.

### Accessory quantities

`NumGaskets` and `NumClamps` are not normal display columns.

They generate accessory BOM rows.

For each grouped pipe row:

```text
TotalGaskets = NumGaskets * PipeQty
TotalClamps = NumClamps * PipeQty
```

Accessory rows should be grouped separately from pipe cut rows.

Example sections:

```text
Pipe Cut List
Fittings
Pipe Accessories
Other Components
```

---

## Persistent mapping profile

Use JSON.

Example profile:

```json
{
  "profileName": "AFCA Pipe BOM",
  "version": 1,
  "partClassRules": [
    {
      "className": "Pipe",
      "detectWhenPropertyExists": "PipeLength"
    }
  ],
  "pipeColumns": [
    {
      "sourceProperty": "BOM",
      "displayName": "BOM Code",
      "enabled": true,
      "groupBy": true,
      "order": 1
    },
    {
      "sourceProperty": "Pipe Identifier",
      "displayName": "Pipe Description",
      "enabled": true,
      "groupBy": true,
      "order": 2
    },
    {
      "sourceProperty": "Specification",
      "displayName": "Specification",
      "enabled": true,
      "groupBy": true,
      "order": 3
    },
    {
      "sourceProperty": "PipeLength",
      "displayName": "Cut Length",
      "enabled": true,
      "groupBy": true,
      "order": 4,
      "unit": "in"
    }
  ],
  "accessoryRules": [
    {
      "sourceProperty": "NumGaskets",
      "displayName": "Gaskets",
      "bomSection": "Pipe Accessories"
    },
    {
      "sourceProperty": "NumClamps",
      "displayName": "Clamps",
      "bomSection": "Pipe Accessories"
    }
  ],
  "ignoredProperties": [
    "BlueGasket",
    "WhiteGasket",
    "BlueFerrule",
    "WhiteFerrule"
  ]
}
```

Storage order:

```text
1. Project-local profile beside the assembly, if present
2. User profile in %AppData%\AFCA\SolidWorksBOMAddin\profiles\
3. Company profile in %ProgramData%\AFCA\SolidWorksBOMAddin\profiles\
4. Built-in default profile
```

---

## Agent 1: repo and solution scaffold

Prompt:

```text
Create the initial C# solution for a SolidWorks BOM add-in.

Requirements:
- Create a solution named SolidWorksBOMAddin.
- Add three projects:
  1. SolidWorksBOMAddin: the SolidWorks COM add-in and UI shell.
  2. BomCore: pure C# domain logic with no SolidWorks references.
  3. BomCore.Tests: unit tests for BomCore.
- Do not put SolidWorks API references in BomCore.
- Add docs/architecture.md explaining the project separation.
- Add profiles/default.pipebom.json using the profile format below.
- Add a README.md with build instructions and the intended workflow.
- Use clean C# naming.
- Do not implement heavy SolidWorks logic yet. Create interfaces and stubs.

Profile format:
[paste the JSON profile above]
```

Acceptance check:

```text
- Solution builds.
- BomCore has no SolidWorks references.
- default.pipebom.json exists.
- README explains that the add-in scans assemblies and generates grouped BOMs.
```

---

## Agent 2: core domain model

Prompt:

```text
Implement the BomCore domain model.

Create these core models:

- ComponentRecord
  - ComponentId
  - FilePath
  - ConfigurationName
  - ComponentName
  - Quantity
  - IsSuppressed
  - IsHidden
  - Properties: dictionary of PropertyValue

- PropertyValue
  - Name
  - RawValue
  - EvaluatedValue
  - Scope
  - Source

- BomProfile
  - ProfileName
  - Version
  - PartClassRules
  - PipeColumns
  - AccessoryRules
  - IgnoredProperties

- BomColumnRule
  - SourceProperty
  - DisplayName
  - Enabled
  - GroupBy
  - Order
  - Unit

- AccessoryRule
  - SourceProperty
  - DisplayName
  - BomSection

- BomRow
  - Section
  - RowType
  - Values dictionary
  - Quantity

- BomResult
  - Rows
  - Diagnostics

Add enums where useful:
- PropertyScope: File, Configuration, Component, CutList, Unknown
- BomRowType: PipeCut, Fitting, Accessory, Other
- DiagnosticSeverity: Info, Warning, Error

Do not add SolidWorks references.

Also add unit tests for:
- loading a BomProfile from JSON
- validating duplicate display names
- validating missing required pipe fields
```

Acceptance check:

```text
- BomCore.Tests pass.
- BomProfile can deserialize default.pipebom.json.
- Invalid profile diagnostics are returned instead of throwing unhandled exceptions.
```

---

## Agent 3: BOM grouping engine

Prompt:

```text
Implement the BOM grouping engine in BomCore.

Create a BomGenerator class that accepts:
- IEnumerable<ComponentRecord>
- BomProfile

It must return BomResult.

Rules:
1. A component is a pipe if it has the property PipeLength.
2. Pipe cut rows are grouped by all enabled pipeColumns where groupBy=true.
3. Pipe cut grouping must include PipeLength.
4. Quantity is the number of matching component instances.
5. Accessory rules generate separate rows.
6. For NumGaskets and NumClamps:
   - parse evaluated value as decimal or integer
   - multiply by grouped pipe quantity
   - skip zero quantities unless profile option ShowZeroAccessoryRows is later added
7. Ignored properties should not appear as normal BOM columns.
8. Non-pipe components should be grouped by BOM if BOM exists, otherwise by file path + configuration.

Add tests using fake ComponentRecord data:
- two identical pipes with same length group into qty 2
- two pipes with same BOM but different PipeLength do not group together
- NumGaskets=2 and pipe qty=4 creates accessory qty 8
- NumClamps=1 and pipe qty=4 creates accessory qty 4
- BlueGasket and WhiteGasket are ignored
- missing PipeLength creates a diagnostic
```

Acceptance check:

```text
- All grouping tests pass.
- Pipes with different lengths stay separated.
- Accessories are in section Pipe Accessories.
```

---

## Agent 4: property discovery and mapping service

Prompt:

```text
Implement a property discovery and mapping service.

In BomCore, create:
- PropertyDiscoveryResult
  - DiscoveredProperties
  - SuggestedColumns
  - IgnoredProperties
  - Diagnostics

Create PropertyDiscoveryService with methods:
- DiscoverFromComponents(IEnumerable<ComponentRecord>)
- SuggestDefaultProfile(IEnumerable<ComponentRecord>)

Rules:
- Build a unique list of all discovered property names.
- Mark BlueGasket, WhiteGasket, BlueFerrule, WhiteFerrule as ignored by default.
- If PipeLength exists, suggest Pipe profile rules.
- If BOM exists, suggest display name BOM Code.
- If Pipe Identifier exists, suggest display name Pipe Description.
- If Specification exists, suggest display name Specification.
- If NumGaskets exists, suggest accessory rule Gaskets.
- If NumClamps exists, suggest accessory rule Clamps.

Add tests.
```

Acceptance check:

```text
- Given the pipe properties from the screenshot, the service suggests:
  - BOM Code
  - Pipe Description
  - Specification
  - Cut Length
  - Gaskets accessory rule
  - Clamps accessory rule
- Ignored properties are marked as ignored, not deleted.
```

---

## Agent 5: SolidWorks API adapter

Prompt:

```text
Implement the SolidWorks adapter layer in the SolidWorksBOMAddin project.

Create interfaces in BomCore:
- IAssemblyReader
  - ReadActiveAssembly(): IReadOnlyList<ComponentRecord>
- ISelectedComponentPropertyReader
  - ReadSelectedComponentProperties(): IReadOnlyList<PropertyValue>

Implement these interfaces in SolidWorksBOMAddin:
- SolidWorksAssemblyReader
- SolidWorksSelectedComponentPropertyReader

Rules:
- Traverse the active assembly recursively.
- Skip suppressed components by default.
- Record suppressed components in diagnostics if possible.
- Read component file path.
- Read referenced configuration.
- Read component name.
- Resolve custom properties in this order:
  1. configuration-specific properties
  2. file-level custom properties
  3. component-level custom properties, if available
- Store both raw and evaluated values.
- Do not crash on lightweight/unresolved components. Add diagnostics.
- Include a method to force-resolve components if the user selects that option later, but do not force it by default.

Do not implement final UI yet. Expose the adapter through testable classes where possible.
```

Acceptance check:

```text
- Add-in can read selected component properties.
- Add-in can scan an open assembly and return ComponentRecord objects.
- Errors on individual components do not stop the full scan.
```

---

## Agent 6: profile persistence

Prompt:

```text
Implement profile persistence.

Create ProfileStore in BomCore or a small infrastructure layer.

Features:
- Load profile from explicit path.
- Save profile to explicit path.
- Load effective profile using this priority:
  1. Project-local profile beside assembly
  2. User profile in %AppData%\AFCA\SolidWorksBOMAddin\profiles\
  3. Company profile in %ProgramData%\AFCA\SolidWorksBOMAddin\profiles\
  4. Built-in default profile
- Validate profile after loading.
- If profile is invalid, return diagnostics and fall back to built-in default.

Use JSON with indentation.

Add tests using temp folders.
```

Acceptance check:

```text
- User profile persists after restart.
- Project-local profile overrides user profile.
- Invalid JSON does not crash the add-in.
```

---

## Agent 7: export engine

Prompt:

```text
Implement export engines.

In BomCore, create:
- IBomExporter
  - Export(BomResult result, Stream output)

Implement:
- CsvBomExporter
- XlsxBomExporter

CSV requirements:
- One section at a time.
- Include section header.
- Include column headers.
- Include Quantity column.
- Use invariant-safe escaping.

XLSX requirements:
- Use a maintained .NET Excel library already compatible with the project.
- One worksheet named BOM.
- Separate sections with blank rows.
- Freeze header row if practical.
- Autosize columns if the library supports it.
- Do not require Microsoft Excel to be installed.

Output sections:
- Pipe Cut List
- Fittings
- Pipe Accessories
- Other Components

Add tests for CSV content.
```

Acceptance check:

```text
- CSV export works without SolidWorks.
- XLSX export works without Excel installed.
- Pipe accessories appear in their own section.
```

---

## Agent 8: add-in UI

Prompt:

```text
Build the first usable UI for the SolidWorks add-in.

UI commands:
1. Read Selected Part Properties
2. Scan Active Assembly
3. Edit BOM Mapping
4. Generate BOM Preview
5. Export CSV
6. Export Excel

UI behavior:
- Read Selected Part Properties opens a property grid showing:
  - Property Name
  - Raw Value
  - Evaluated Value
  - Scope
  - Include checkbox
  - Display Name text field
  - Group By checkbox
  - BOM Section
- Scan Active Assembly discovers all properties from all components.
- Edit BOM Mapping loads current effective profile.
- Generate BOM Preview shows grouped rows before export.
- Export CSV and Export Excel ask for save path.

Do not overdesign the UI.
A simple WinForms or WPF task pane is acceptable.
Prefer reliability over visual polish.

The UI must make it clear that NumGaskets and NumClamps generate accessory rows, not normal pipe columns.
```

Acceptance check:

```text
- User can select a pipe part and see the properties from the screenshot.
- User can choose BOM, Pipe Identifier, Specification, PipeLength.
- User can rename them to BOM Code, Pipe Description, Specification, Cut Length.
- User can save the profile.
- User can scan the full assembly and export a BOM.
```

---

## Agent 9: diagnostics and debug report

Prompt:

```text
Add diagnostics and a debug report.

Diagnostics must include:
- Components skipped because suppressed
- Components skipped because unresolved
- Components missing BOM
- Pipes missing PipeLength
- Pipes missing Pipe Identifier
- Pipes missing Specification
- Accessory properties that could not be parsed as numbers
- Duplicate display names in the profile
- Components with no readable model document

Create a DebugReportExporter that writes:
- assembly path
- profile path used
- number of components scanned
- number of components skipped
- all unique discovered properties
- generated BOM row count
- diagnostics list

Export as .txt or .json.
```

Acceptance check:

```text
- A bad component produces a warning, not a crash.
- User can export a debug report and send it for troubleshooting.
```

---

## Agent 10: installer and registration

Prompt:

```text
Prepare the add-in for installation and registration.

Requirements:
- Document how to build the add-in.
- Document how to register/unregister the COM add-in.
- Add scripts if appropriate:
  - register-addin.ps1
  - unregister-addin.ps1
- Add clear notes about required SolidWorks version, .NET runtime, and admin permissions.
- Do not hardcode local developer paths.
- Make the add-in name visible in SolidWorks Add-Ins list as:
  AFCA Piping BOM Generator

Also document where profiles are stored:
- %AppData%\AFCA\SolidWorksBOMAddin\profiles\
- %ProgramData%\AFCA\SolidWorksBOMAddin\profiles\
```

Acceptance check:

```text
- Fresh machine instructions are clear.
- Add-in can be registered and appears in SolidWorks.
- Add-in can be unregistered cleanly.
```

---

## Implementation order

Do it in this order:

```text
1. Repo scaffold
2. BomCore models
3. Profile JSON loading
4. BOM grouping tests
5. Property discovery tests
6. CSV export
7. XLSX export
8. SolidWorks selected-part property reader
9. SolidWorks assembly scanner
10. Minimal UI
11. Diagnostics report
12. Installer/registration cleanup
```

Do not build the full UI first. The highest-risk logic is the assembly scanning and grouping, not the buttons.

---

## Master prompt for Copilot CLI

Use this as the first high-level instruction:

```text
We are building a SolidWorks C# add-in called AFCA Piping BOM Generator.

The add-in scans a SolidWorks piping assembly, reads every component's properties, groups parts into a BOM, and exports CSV/XLSX.

Keep the architecture split:
- BomCore: pure C# logic, no SolidWorks references.
- SolidWorksBOMAddin: COM add-in, SolidWorks API adapter, UI.
- BomCore.Tests: tests for profiles, grouping, property discovery, and exporters.

Known pipe properties:
- Pipe Identifier
- SWbompartno
- Specification
- BlueGasket
- WhiteGasket
- NumGaskets
- PipeLength
- NumClamps
- Date
- Author
- WhiteFerrule
- BlueFerrule
- BOM

Rules:
- A component is a pipe if it has PipeLength.
- Pipe cut rows are grouped by BOM + Pipe Identifier + Specification + PipeLength.
- NumGaskets and NumClamps generate separate accessory BOM rows.
- Ignore BlueGasket, WhiteGasket, WhiteFerrule, BlueFerrule by default.
- The user can select a part, read its properties, choose properties to include, rename the column headers, and save that mapping persistently as JSON.
- Mapping profiles are loaded in priority order:
  1. project-local profile
  2. user AppData profile
  3. ProgramData company profile
  4. built-in default profile

Do not put SolidWorks API references inside BomCore.
Write tests for all core logic before building the UI.
Prefer explicit diagnostics over crashes.
```

---

## Definition of done for first working version

The first useful version is done when this works:

```text
1. Open SolidWorks assembly.
2. Select one pipe.
3. Click Read Selected Part Properties.
4. See Pipe Identifier, Specification, NumGaskets, PipeLength, NumClamps, BOM.
5. Map:
   - BOM → BOM Code
   - Pipe Identifier → Pipe Description
   - Specification → Specification
   - PipeLength → Cut Length
6. Save profile.
7. Scan full assembly.
8. Generate BOM preview.
9. Export Excel.
10. Pipe cut list separates same BOM codes with different lengths.
11. Gaskets and clamps appear in Pipe Accessories.
```

That is the minimum version worth testing on real assemblies.
