namespace BomCore;

public sealed class BomGenerator
{
    public BomResult Generate(IEnumerable<ComponentRecord> components, BomProfile profile)
    {
        var componentList = components.ToList();
        var diagnostics = new List<BomDiagnostic>(BomProfileSerializer.Validate(profile));
        var rows = new List<BomRow>();
        var pipeDetectorProperty = profile.GetPipeDetectionProperty();
        var enabledPipeColumns = profile.PipeColumns
            .Where(column => column.Enabled)
            .OrderBy(column => column.Order)
            .ToList();
        var ignoredProperties = new HashSet<string>(profile.IgnoredProperties, StringComparer.OrdinalIgnoreCase);

        foreach (var component in componentList)
        {
            if (component.IsSuppressed)
            {
                diagnostics.Add(new BomDiagnostic
                {
                    Severity = DiagnosticSeverity.Info,
                    Code = "component-suppressed",
                    Message = $"Component '{component.ComponentName}' was skipped because it is suppressed.",
                    ComponentId = component.ComponentId,
                });
                continue;
            }

            if (LooksLikePipeCandidate(component, pipeDetectorProperty) && component.GetPropertyValue(pipeDetectorProperty) is null)
            {
                diagnostics.Add(new BomDiagnostic
                {
                    Severity = DiagnosticSeverity.Warning,
                    Code = "missing-pipe-length",
                    Message = $"Component '{component.ComponentName}' appears to be a pipe candidate but is missing PipeLength.",
                    ComponentId = component.ComponentId,
                    PropertyName = pipeDetectorProperty,
                });
            }

            if (IsPipe(component, pipeDetectorProperty))
            {
                AddMissingPipeFieldDiagnostics(component, diagnostics);
            }
            else if (string.IsNullOrWhiteSpace(component.GetPropertyValue(KnownPropertyNames.Bom)))
            {
                diagnostics.Add(new BomDiagnostic
                {
                    Severity = DiagnosticSeverity.Info,
                    Code = "missing-bom",
                    Message = $"Component '{component.ComponentName}' does not define BOM.",
                    ComponentId = component.ComponentId,
                    PropertyName = KnownPropertyNames.Bom,
                });
            }
        }

        var activeComponents = componentList.Where(component => !component.IsSuppressed).ToList();

        var pipeGroups = activeComponents
            .Where(component => IsPipe(component, pipeDetectorProperty))
            .GroupBy(component => BuildPipeGroupKey(component, enabledPipeColumns));

        foreach (var group in pipeGroups)
        {
            var groupedComponents = group.ToList();
            var firstComponent = groupedComponents[0];
            var quantity = groupedComponents.Sum(component => component.Quantity);
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (var column in enabledPipeColumns)
            {
                if (ignoredProperties.Contains(column.SourceProperty))
                {
                    continue;
                }

                values[column.DisplayName] = firstComponent.GetPropertyValue(column.SourceProperty) ?? string.Empty;
            }

            rows.Add(new BomRow
            {
                Section = KnownBomSections.PipeCutList,
                RowType = BomRowType.PipeCut,
                Values = values,
                Quantity = quantity,
            });

            foreach (var accessoryRule in profile.AccessoryRules)
            {
                decimal accessoryQuantity = 0m;

                foreach (var component in groupedComponents)
                {
                    var rawValue = component.GetPropertyValue(accessoryRule.SourceProperty);
                    if (string.IsNullOrWhiteSpace(rawValue))
                    {
                        continue;
                    }

                    if (!NumericParsing.TryParseDecimal(rawValue, out var numericValue))
                    {
                        diagnostics.Add(new BomDiagnostic
                        {
                            Severity = DiagnosticSeverity.Warning,
                            Code = "invalid-accessory-quantity",
                            Message = $"Property '{accessoryRule.SourceProperty}' on component '{component.ComponentName}' could not be parsed as a number.",
                            ComponentId = component.ComponentId,
                            PropertyName = accessoryRule.SourceProperty,
                        });
                        continue;
                    }

                    accessoryQuantity += numericValue * component.Quantity;
                }

                if (accessoryQuantity == 0m)
                {
                    continue;
                }

                rows.Add(new BomRow
                {
                    Section = string.IsNullOrWhiteSpace(accessoryRule.BomSection) ? KnownBomSections.PipeAccessories : accessoryRule.BomSection,
                    RowType = BomRowType.Accessory,
                    Values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["Description"] = accessoryRule.DisplayName,
                    },
                    Quantity = accessoryQuantity,
                });
            }
        }

        var nonPipeGroups = activeComponents
            .Where(component => !IsPipe(component, pipeDetectorProperty))
            .GroupBy(BuildNonPipeGroupKey);

        foreach (var group in nonPipeGroups)
        {
            var groupedComponents = group.ToList();
            var firstComponent = groupedComponents[0];
            var quantity = groupedComponents.Sum(component => component.Quantity);
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var bomValue = firstComponent.GetPropertyValue(KnownPropertyNames.Bom);

            if (!string.IsNullOrWhiteSpace(bomValue))
            {
                values["BOM"] = bomValue;
            }

            values["Component"] = firstComponent.ComponentName;

            if (firstComponent.IsVirtual)
            {
                values["Component Type"] = "Virtual";
            }
            else if (!string.IsNullOrWhiteSpace(firstComponent.FilePath))
            {
                values["File Path"] = firstComponent.FilePath!;
            }

            if (!string.IsNullOrWhiteSpace(firstComponent.ConfigurationName))
            {
                values["Configuration"] = firstComponent.ConfigurationName;
            }

            rows.Add(new BomRow
            {
                Section = KnownBomSections.OtherComponents,
                RowType = BomRowType.Other,
                Values = values,
                Quantity = quantity,
            });
        }

        return new BomResult
        {
            Rows = OrderRows(rows),
            Diagnostics = diagnostics,
        };
    }

    private static IReadOnlyList<BomRow> OrderRows(IEnumerable<BomRow> rows)
    {
        var sectionOrder = KnownBomSections.DisplayOrder
            .Select((section, index) => new { section, index })
            .ToDictionary(entry => entry.section, entry => entry.index, StringComparer.OrdinalIgnoreCase);

        return rows
            .OrderBy(row => sectionOrder.GetValueOrDefault(row.Section, int.MaxValue))
            .ThenBy(row => row.RowType)
            .ThenBy(row => string.Join("|", row.Values.Values))
            .ToList();
    }

    private static bool IsPipe(ComponentRecord component, string pipeDetectorProperty)
    {
        return !string.IsNullOrWhiteSpace(component.GetPropertyValue(pipeDetectorProperty));
    }

    private static bool LooksLikePipeCandidate(ComponentRecord component, string pipeDetectorProperty)
    {
        if (component.GetPropertyValue(pipeDetectorProperty) is not null)
        {
            return true;
        }

        return KnownPropertyNames.PipeCandidateProperties.Any(propertyName => component.GetPropertyValue(propertyName) is not null);
    }

    private static string BuildPipeGroupKey(ComponentRecord component, IEnumerable<BomColumnRule> enabledPipeColumns)
    {
        return string.Join(
            "\u001F",
            enabledPipeColumns
                .Where(column => column.GroupBy)
                .Select(column => component.GetPropertyValue(column.SourceProperty) ?? string.Empty));
    }

    private static string BuildNonPipeGroupKey(ComponentRecord component)
    {
        var bomValue = component.GetPropertyValue(KnownPropertyNames.Bom);
        if (!string.IsNullOrWhiteSpace(bomValue))
        {
            return $"bom:{bomValue}";
        }

        return component.GetIdentityFallback();
    }

    private static void AddMissingPipeFieldDiagnostics(ComponentRecord component, ICollection<BomDiagnostic> diagnostics)
    {
        foreach (var propertyName in new[]
                 {
                     KnownPropertyNames.Bom,
                     KnownPropertyNames.PipeIdentifier,
                     KnownPropertyNames.Specification,
                 })
        {
            if (component.GetPropertyValue(propertyName) is not null)
            {
                continue;
            }

            diagnostics.Add(new BomDiagnostic
            {
                Severity = DiagnosticSeverity.Warning,
                Code = $"missing-{propertyName.ToLowerInvariant().Replace(' ', '-')}",
                Message = $"Pipe component '{component.ComponentName}' is missing '{propertyName}'.",
                ComponentId = component.ComponentId,
                PropertyName = propertyName,
            });
        }
    }
}
