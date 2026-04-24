namespace BomCore;

public sealed class DebugReportService
{
    private readonly PropertyDiscoveryService _propertyDiscoveryService;

    public DebugReportService(PropertyDiscoveryService? propertyDiscoveryService = null)
    {
        _propertyDiscoveryService = propertyDiscoveryService ?? new PropertyDiscoveryService();
    }

    public DebugReport Create(DebugReportInput input)
    {
        ArgumentNullException.ThrowIfNull(input);

        var discoveredProperties = _propertyDiscoveryService
            .DiscoverFromComponents(input.Components)
            .DiscoveredProperties;

        return new DebugReport
        {
            AssemblyPath = input.AssemblyPath,
            ProfilePath = input.ProfilePath,
            ComponentsScanned = input.ComponentsScanned ?? input.Components.Count,
            ComponentsSkipped = input.ComponentsSkipped,
            DiscoveredProperties = discoveredProperties,
            GeneratedBomRowCount = input.Rows.Count,
            Diagnostics = input.Diagnostics,
        };
    }
}
