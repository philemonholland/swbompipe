using BomCore;
using SolidWorks.Interop.sldworks;

namespace SolidWorksBOMAddin;

public sealed class SolidWorksSelectedComponentPropertyReader : ISelectedComponentPropertyReader
{
    private readonly ISldWorks _application;

    public SolidWorksSelectedComponentPropertyReader(ISldWorks application)
    {
        _application = application ?? throw new ArgumentNullException(nameof(application));
    }

    public IReadOnlyList<PropertyValue> ReadSelectedComponentProperties()
    {
        var modelDocument = _application.IActiveDoc2
            ?? throw new InvalidOperationException("No active SolidWorks document is open.");
        var selectionManager = modelDocument.ISelectionManager;
        if (selectionManager is null || selectionManager.GetSelectedObjectCount2(-1) < 1)
        {
            return [];
        }

        var component = selectionManager.GetSelectedObjectsComponent3(1, -1) as IComponent2;
        if (component is not null)
        {
            var referencedModel = component.GetModelDoc2() as IModelDoc2;
            return SolidWorksPropertyExtractor
                .ReadProperties(component, referencedModel, component.ReferencedConfiguration)
                .Values
                .ToList();
        }

        var selectedDocument = selectionManager.GetSelectedObject6(1, -1) as IModelDoc2;
        if (selectedDocument is null)
        {
            return [];
        }

        var activeConfiguration = selectedDocument.ConfigurationManager?.ActiveConfiguration;
        return SolidWorksPropertyExtractor
            .ReadProperties(component: null, selectedDocument, activeConfiguration?.Name)
            .Values
            .ToList();
    }
}
