# Property and grouping rules

## Pipe detection

A component is treated as a pipe when the property `PipeLength` exists. Pipe handling is therefore property-driven rather than file-name-driven.

Known pipe-related properties from the planning brief:

- `Pipe Identifier`
- `SWbompartno`
- `Specification`
- `BlueGasket`
- `WhiteGasket`
- `NumGaskets`
- `PipeLength`
- `NumClamps`
- `Date`
- `Author`
- `WhiteFerrule`
- `BlueFerrule`
- `BOM`

## Ignored-by-default properties

These properties stay discoverable in the property explorer, but should not be selected as normal BOM columns by default:

- `BlueGasket`
- `WhiteGasket`
- `BlueFerrule`
- `WhiteFerrule`

## Pipe grouping rules

Pipe cut rows must be grouped by all of the following values:

- `BOM`
- `Pipe Identifier`
- `Specification`
- `PipeLength`

`PipeLength` is mandatory for grouping. Pipes with the same BOM code but different cut lengths must remain separate rows.

## Accessory rules

`NumGaskets` and `NumClamps` are accessory generators, not normal display columns.

For each grouped pipe row:

- `TotalGaskets = NumGaskets * PipeQty`
- `TotalClamps = NumClamps * PipeQty`

Accessory rows should be emitted in a separate **Pipe Accessories** section instead of being merged into the pipe cut list. Zero or blank accessory quantities should be skipped unless a later profile option explicitly enables zero-row output.

## Virtual parts and subassemblies

- Assembly traversal should recurse through subassemblies and evaluate component instances, not just top-level nodes.
- BOM grouping should be based on the effective properties of each discovered component, not on which subassembly contains it.
- Suppressed or unresolved components should be skipped by default and surfaced through diagnostics instead of causing a hard failure.
- Virtual parts may not have a stable external file path; the adapter should preserve whatever component identity and readable properties are available, and record diagnostics if the model document or path cannot be resolved cleanly.

## Diagnostics-first behavior

Prefer explicit diagnostics over crashes for rule violations and bad input. Examples include missing `PipeLength`, unreadable component data, or accessory values that cannot be parsed as numbers.
