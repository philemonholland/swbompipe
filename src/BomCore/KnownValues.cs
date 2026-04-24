namespace BomCore;

public static class KnownPropertyNames
{
    public const string Bom = "BOM";
    public const string PipeIdentifier = "Pipe Identifier";
    public const string Specification = "Specification";
    public const string PipeLength = "PipeLength";
    public const string NumGaskets = "NumGaskets";
    public const string NumClamps = "NumClamps";
    public const string BlueGasket = "BlueGasket";
    public const string WhiteGasket = "WhiteGasket";
    public const string BlueFerrule = "BlueFerrule";
    public const string WhiteFerrule = "WhiteFerrule";

    public static readonly IReadOnlyList<string> DefaultIgnoredProperties =
    [
        BlueGasket,
        WhiteGasket,
        BlueFerrule,
        WhiteFerrule,
    ];

    public static readonly IReadOnlyList<string> PipeRequiredProperties =
    [
        Bom,
        PipeIdentifier,
        Specification,
        PipeLength,
    ];

    public static readonly IReadOnlyList<string> PipeCandidateProperties =
    [
        Bom,
        PipeIdentifier,
        Specification,
        NumGaskets,
        NumClamps,
    ];
}

public static class KnownBomSections
{
    public const string PipeCutList = "Pipe Cut List";
    public const string Fittings = "Fittings";
    public const string PipeAccessories = "Pipe Accessories";
    public const string OtherComponents = "Other Components";

    public static readonly IReadOnlyList<string> DisplayOrder =
    [
        PipeCutList,
        Fittings,
        PipeAccessories,
        OtherComponents,
    ];
}
