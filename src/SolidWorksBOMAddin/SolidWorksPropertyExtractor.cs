using BomCore;
using SolidWorks.Interop.sldworks;

namespace SolidWorksBOMAddin;

internal static class SolidWorksPropertyExtractor
{
    public static IReadOnlyDictionary<string, PropertyValue> ReadProperties(
        IComponent2? component,
        IModelDoc2? modelDocument,
        string? configurationName)
    {
        var properties = new Dictionary<string, PropertyValue>(StringComparer.OrdinalIgnoreCase);
        if (modelDocument is null)
        {
            return properties;
        }

        AddProperties(
            properties,
            GetManager(modelDocument, configurationName),
            PropertyScope.Configuration,
            configurationName);

        AddProperties(
            properties,
            GetManager(modelDocument, string.Empty),
            PropertyScope.File,
            modelDocument.GetPathName());

        if (component is not null)
        {
            AddProperties(
                properties,
                component.CustomPropertyManager[configurationName ?? string.Empty] as ICustomPropertyManager,
                PropertyScope.Component,
                component.Name2);
        }

        return properties;
    }

    private static ICustomPropertyManager? GetManager(IModelDoc2 modelDocument, string? configurationName)
    {
        return modelDocument.Extension.CustomPropertyManager[configurationName ?? string.Empty];
    }

    private static void AddProperties(
        IDictionary<string, PropertyValue> properties,
        ICustomPropertyManager? propertyManager,
        PropertyScope scope,
        string? source)
    {
        if (propertyManager is null)
        {
            return;
        }

        foreach (var propertyName in EnumeratePropertyNames(propertyManager))
        {
            propertyManager.Get6(
                propertyName,
                UseCached: false,
                out var rawValue,
                out var resolvedValue,
                out _,
                out _);

            AddOrReplaceBlankValue(
                properties,
                propertyName,
                new PropertyValue
                {
                    Name = propertyName,
                    RawValue = rawValue,
                    EvaluatedValue = resolvedValue,
                    Scope = scope,
                    Source = source,
                });
        }
    }

    private static void AddOrReplaceBlankValue(
        IDictionary<string, PropertyValue> properties,
        string propertyName,
        PropertyValue propertyValue)
    {
        if (!properties.TryGetValue(propertyName, out var existingValue))
        {
            properties[propertyName] = propertyValue;
            return;
        }

        if (string.IsNullOrWhiteSpace(existingValue.EffectiveValue)
            && !string.IsNullOrWhiteSpace(propertyValue.EffectiveValue))
        {
            properties[propertyName] = propertyValue;
        }
    }

    private static IEnumerable<string> EnumeratePropertyNames(ICustomPropertyManager propertyManager)
    {
        var names = propertyManager.GetNames();
        return names switch
        {
            string[] stringArray => stringArray,
            object[] objectArray => objectArray.OfType<string>(),
            string singleName => [singleName],
            _ => [],
        };
    }
}
