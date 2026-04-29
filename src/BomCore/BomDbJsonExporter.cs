using System.Text.Json;
using System.Text.Json.Serialization;

namespace BomCore;

public sealed class BomDbJsonExporter
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    };

    public void Export(BomDbImportFile payload, Stream output)
    {
        ArgumentNullException.ThrowIfNull(payload);
        ArgumentNullException.ThrowIfNull(output);

        JsonSerializer.Serialize(output, payload, JsonOptions);
    }
}
