namespace BomCore;

public sealed class BomGenerator
{
    public BomResult Generate(IEnumerable<ComponentRecord> components, BomProfile profile)
    {
        var componentList = components.ToList();
        var diagnostics = new List<BomDiagnostic>(BomProfileSerializer.Validate(profile));
        var rows = new List<BomRow>();
        var ignoredProperties = new HashSet<string>(profile.IgnoredProperties, StringComparer.OrdinalIgnoreCase);
        var activeComponents = componentList.Where(component => !component.IsSuppressed).ToList();

        foreach (var component in componentList.Where(component => component.IsSuppressed))
        {
            diagnostics.Add(new BomDiagnostic
            {
                Severity = DiagnosticSeverity.Info,
                Code = "component-suppressed",
                Message = $"Component '{component.ComponentName}' was skipped because it is suppressed.",
                ComponentId = component.ComponentId,
            });
        }

        var classifiedComponents = activeComponents
            .Select(component => new ClassifiedComponent(
                component,
                ClassifyComponent(component, profile.GetEffectiveSectionRules(), diagnostics)))
            .ToList();

        foreach (var sectionGroup in classifiedComponents.GroupBy(item => item.Section, StringComparer.OrdinalIgnoreCase))
        {
            var section = KnownBomSections.NormalizeConfigurableSection(sectionGroup.Key);
            var columns = profile.GetSectionColumns(section)
                .Where(column => column.Enabled && !ignoredProperties.Contains(column.SourceProperty))
                .OrderBy(column => column.Order)
                .ToList();

            foreach (var group in sectionGroup.GroupBy(item => BuildGroupKey(item.Component, columns)))
            {
                var groupedComponents = group.Select(item => item.Component).ToList();
                var quantity = groupedComponents.Sum(component => component.Quantity);
                var firstComponent = groupedComponents[0];
                var values = BuildValues(firstComponent, columns, diagnostics);

                rows.Add(new BomRow
                {
                    Section = section,
                    RowType = GetRowType(section),
                    Values = values,
                    Quantity = quantity,
                });

                if (string.Equals(section, KnownBomSections.Pipes, StringComparison.OrdinalIgnoreCase))
                {
                    AddAccessoryRows(groupedComponents, profile.AccessoryRules, rows, diagnostics);
                }
            }
        }

        return new BomResult
        {
            Rows = OrderRows(MergeDuplicateAccessoryRows(rows)),
            Diagnostics = diagnostics,
        };
    }

    private static string ClassifyComponent(
        ComponentRecord component,
        IReadOnlyList<BomSectionRule> sectionRules,
        ICollection<BomDiagnostic> diagnostics)
    {
        var matchedRule = sectionRules.FirstOrDefault(rule =>
            string.Equals(component.GetPropertyValue(rule.SourceProperty), rule.MatchValue, StringComparison.OrdinalIgnoreCase));
        if (matchedRule is not null)
        {
            return KnownBomSections.NormalizeConfigurableSection(matchedRule.Section);
        }

        var familyValue = component.GetPropertyValue(KnownPropertyNames.PrimaryFamily);
        if (string.IsNullOrWhiteSpace(familyValue))
        {
            return KnownBomSections.Other;
        }

        var aliasSection = KnownBomSections.DisplayOrder
            .Concat(KnownBomSections.PreferredDynamicSections)
            .FirstOrDefault(section =>
                string.Equals(familyValue, section, StringComparison.OrdinalIgnoreCase)
                || string.Equals(familyValue, ToSingularAlias(section), StringComparison.OrdinalIgnoreCase));

        if (aliasSection is not null)
        {
            return aliasSection;
        }

        return KnownBomSections.NormalizeConfigurableSection(familyValue);
    }

    private static IReadOnlyDictionary<string, string> BuildValues(
        ComponentRecord component,
        IEnumerable<BomColumnRule> columns,
        ICollection<BomDiagnostic> diagnostics)
    {
        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var column in columns)
        {
            if (string.IsNullOrWhiteSpace(column.SourceProperty) || string.IsNullOrWhiteSpace(column.DisplayName))
            {
                continue;
            }

            var value = component.GetPropertyValue(column.SourceProperty);
            if (value is null)
            {
                diagnostics.Add(new BomDiagnostic
                {
                    Severity = DiagnosticSeverity.Warning,
                    Code = "missing-configured-property",
                    Message = $"Configured property '{column.SourceProperty}' is missing on component '{component.ComponentName}'.",
                    ComponentId = component.ComponentId,
                    PropertyName = column.SourceProperty,
                });
            }

            values[column.DisplayName] = value ?? string.Empty;
        }

        return values;
    }

    private static string BuildGroupKey(ComponentRecord component, IEnumerable<BomColumnRule> columns)
    {
        var groupByColumns = columns.Where(column => column.GroupBy).ToList();
        if (groupByColumns.Count == 0)
        {
            groupByColumns = columns.ToList();
        }

        return string.Join(
            "\u001F",
            groupByColumns.Select(column => component.GetPropertyValue(column.SourceProperty) ?? string.Empty));
    }

    private static string ToSingularAlias(string section)
    {
        return section.EndsWith("s", StringComparison.OrdinalIgnoreCase)
            ? section[..^1]
            : section;
    }

    private static void AddAccessoryRows(
        IReadOnlyList<ComponentRecord> groupedComponents,
        IEnumerable<AccessoryRule> accessoryRules,
        ICollection<BomRow> rows,
        ICollection<BomDiagnostic> diagnostics)
    {
        foreach (var accessoryRule in accessoryRules)
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
                Section = KnownBomSections.NormalizeAccessorySection(accessoryRule.BomSection),
                RowType = BomRowType.Accessory,
                Values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["Description"] = accessoryRule.DisplayName,
                },
                Quantity = accessoryQuantity,
            });
        }
    }

    private static IReadOnlyList<BomRow> OrderRows(IEnumerable<BomRow> rows)
    {
        var orderedSections = KnownBomSections.OrderSections(rows.Select(row => row.Section));
        var sectionOrder = orderedSections
            .Select((section, index) => new { section, index })
            .ToDictionary(entry => entry.section, entry => entry.index, StringComparer.OrdinalIgnoreCase);

        return rows
            .OrderBy(row => sectionOrder.GetValueOrDefault(row.Section, int.MaxValue))
            .ThenBy(row => row.RowType)
            .ThenBy(row => string.Join("|", row.Values.Values))
            .ToList();
    }

    private static IReadOnlyList<BomRow> MergeDuplicateAccessoryRows(IEnumerable<BomRow> rows)
    {
        var rowList = rows.ToList();
        var mergedRows = rowList
            .Where(row => row.RowType != BomRowType.Accessory)
            .ToList();

        foreach (var group in rowList
                     .Where(row => row.RowType == BomRowType.Accessory)
                     .GroupBy(BuildAccessoryMergeKey, StringComparer.Ordinal))
        {
            var firstRow = group.First();
            mergedRows.Add(new BomRow
            {
                Section = firstRow.Section,
                RowType = firstRow.RowType,
                Values = new Dictionary<string, string>(firstRow.Values, StringComparer.OrdinalIgnoreCase),
                Quantity = group.Sum(row => row.Quantity),
            });
        }

        return mergedRows;
    }

    private static string BuildAccessoryMergeKey(BomRow row)
    {
        var valueKey = string.Join(
            "\u001F",
            row.Values
                .OrderBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase)
                .Select(pair => $"{pair.Key}={pair.Value}"));

        return $"{row.Section}\u001E{valueKey}";
    }

    private static BomRowType GetRowType(string section)
    {
        if (string.Equals(section, KnownBomSections.Pipes, StringComparison.OrdinalIgnoreCase))
        {
            return BomRowType.PipeCut;
        }

        if (string.Equals(section, KnownBomSections.Fittings, StringComparison.OrdinalIgnoreCase))
        {
            return BomRowType.Fitting;
        }

        return BomRowType.SectionItem;
    }

    private sealed record ClassifiedComponent(ComponentRecord Component, string Section);
}
