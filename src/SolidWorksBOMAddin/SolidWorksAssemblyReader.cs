using BomCore;
using SolidWorks.Interop.sldworks;

namespace SolidWorksBOMAddin;

public sealed class SolidWorksAssemblyReader : IAssemblyReader
{
    private readonly ISldWorks _application;
    private readonly SolidWorksReaderOptions _options;

    public SolidWorksAssemblyReader(ISldWorks application, SolidWorksReaderOptions? options = null)
    {
        _application = application ?? throw new ArgumentNullException(nameof(application));
        _options = options ?? new SolidWorksReaderOptions();
    }

    public AssemblyReadResult ReadActiveAssembly()
    {
        var activeDocument = _application.IActiveDoc2
            ?? throw new InvalidOperationException("No active SolidWorks document is open.");
        var assemblyDocument = activeDocument as IAssemblyDoc
            ?? throw new InvalidOperationException("The active SolidWorks document is not an assembly.");

        if (_options.ResolveLightweightComponents)
        {
            assemblyDocument.ResolveAllLightWeightComponents(false);
        }

        var activeConfiguration = activeDocument.ConfigurationManager?.ActiveConfiguration
            ?? throw new InvalidOperationException("The active assembly has no active configuration.");
        var rootComponent = activeConfiguration.GetRootComponent3(_options.ResolveLightweightComponents);
        if (rootComponent is null)
        {
            return new AssemblyReadResult
            {
                AssemblyPath = activeDocument.GetPathName(),
                Components = [],
                ComponentsScanned = 0,
                ComponentsSkipped = 0,
                Diagnostics = [],
            };
        }

        var components = new List<ComponentRecord>();
        var diagnostics = new List<BomDiagnostic>();
        var scannedCount = 0;
        var skippedCount = 0;

        foreach (var child in EnumerateChildren(rootComponent))
        {
            Traverse(child, parentComponentId: null, components, diagnostics, ref scannedCount, ref skippedCount);
        }

        return new AssemblyReadResult
        {
            AssemblyPath = activeDocument.GetPathName(),
            Components = components,
            ComponentsScanned = scannedCount,
            ComponentsSkipped = skippedCount,
            Diagnostics = diagnostics,
        };
    }

    private void Traverse(
        IComponent2 component,
        string? parentComponentId,
        ICollection<ComponentRecord> components,
        ICollection<BomDiagnostic> diagnostics,
        ref int scannedCount,
        ref int skippedCount)
    {
        scannedCount++;

        var componentId = BuildComponentId(component, parentComponentId);
        var componentName = component.Name2 ?? component.GetPathName() ?? componentId;
        var isSuppressed = component.IsSuppressed();
        if (isSuppressed && !_options.IncludeSuppressedComponents)
        {
            diagnostics.Add(new BomDiagnostic
            {
                Severity = DiagnosticSeverity.Info,
                Code = "component-suppressed",
                Message = $"Component '{componentName}' was skipped because it is suppressed.",
                ComponentId = componentId,
            });
            skippedCount++;
            return;
        }

        var childComponents = EnumerateChildren(component).ToList();
        var modelDocument = component.GetModelDoc2() as IModelDoc2;
        if (modelDocument is null)
        {
            diagnostics.Add(new BomDiagnostic
            {
                Severity = DiagnosticSeverity.Warning,
                Code = "component-unresolved",
                Message = $"Component '{componentName}' was skipped because its model could not be resolved.",
                ComponentId = componentId,
            });
            diagnostics.Add(new BomDiagnostic
            {
                Severity = DiagnosticSeverity.Warning,
                Code = "component-no-readable-model",
                Message = $"Component '{componentName}' has no readable model document.",
                ComponentId = componentId,
            });
            skippedCount++;

            foreach (var child in childComponents)
            {
                Traverse(child, componentId, components, diagnostics, ref scannedCount, ref skippedCount);
            }

            return;
        }

        var record = new ComponentRecord
        {
            ComponentId = componentId,
            ParentComponentId = parentComponentId,
            FilePath = component.GetPathName(),
            ConfigurationName = component.ReferencedConfiguration ?? string.Empty,
            ComponentName = componentName,
            Quantity = 1m,
            IsSuppressed = isSuppressed,
            IsHidden = component.IsHidden(ConsiderSuppressed: false),
            IsVirtual = component.IsVirtual,
            IsAssembly = childComponents.Count > 0,
            Properties = SolidWorksPropertyExtractor.ReadProperties(component, modelDocument, component.ReferencedConfiguration),
        };

        components.Add(record);

        foreach (var child in childComponents)
        {
            Traverse(child, record.ComponentId, components, diagnostics, ref scannedCount, ref skippedCount);
        }
    }

    private static string BuildComponentId(IComponent2 component, string? parentComponentId)
    {
        var baseId = component.Name2 ?? component.GetPathName() ?? Guid.NewGuid().ToString("N");
        return string.IsNullOrWhiteSpace(parentComponentId) ? baseId : $"{parentComponentId}/{baseId}";
    }

    private static IEnumerable<IComponent2> EnumerateChildren(IComponent2 component)
    {
        var children = component.GetChildren();
        return children switch
        {
            object[] childArray => childArray.OfType<IComponent2>(),
            IComponent2 singleComponent => [singleComponent],
            _ => [],
        };
    }
}
