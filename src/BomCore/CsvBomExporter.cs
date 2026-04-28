using System.Globalization;
using System.Text;

namespace BomCore;

public sealed class CsvBomExporter : IBomExporter
{
    public void Export(BomResult result, Stream output)
    {
        ArgumentNullException.ThrowIfNull(result);
        ArgumentNullException.ThrowIfNull(output);

        using var writer = new StreamWriter(output, new UTF8Encoding(false), leaveOpen: true);

        foreach (var section in KnownBomSections.OrderSections(result.Rows.Select(row => row.Section)))
        {
            var sectionRows = result.Rows
                .Where(row => string.Equals(row.Section, section, StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (sectionRows.Count == 0)
            {
                continue;
            }

            writer.WriteLine(Escape(section));
            var headers = CollectHeaders(sectionRows);
            writer.WriteLine(string.Join(",", headers.Concat(["Quantity"]).Select(Escape)));

            foreach (var row in sectionRows)
            {
                var values = headers
                    .Select(header => row.Values.TryGetValue(header, out var value) ? value : string.Empty)
                    .Append(row.Quantity.ToString(CultureInfo.InvariantCulture));

                writer.WriteLine(string.Join(",", values.Select(Escape)));
            }

            writer.WriteLine();
        }
    }

    private static IReadOnlyList<string> CollectHeaders(IEnumerable<BomRow> rows)
    {
        var headers = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var row in rows)
        {
            foreach (var header in row.Values.Keys)
            {
                if (seen.Add(header))
                {
                    headers.Add(header);
                }
            }
        }

        return headers;
    }

    private static string Escape(string value)
    {
        var normalized = value.Replace("\"", "\"\"");
        return normalized.IndexOfAny([',', '"', '\r', '\n']) >= 0
            ? $"\"{normalized}\""
            : normalized;
    }
}
