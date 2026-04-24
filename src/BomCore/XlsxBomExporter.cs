using ClosedXML.Excel;

namespace BomCore;

public sealed class XlsxBomExporter : IBomExporter
{
    public void Export(BomResult result, Stream output)
    {
        ArgumentNullException.ThrowIfNull(result);
        ArgumentNullException.ThrowIfNull(output);

        using var workbook = new XLWorkbook();
        var worksheet = workbook.AddWorksheet("BOM");
        var currentRow = 1;

        foreach (var section in KnownBomSections.DisplayOrder)
        {
            var sectionRows = result.Rows
                .Where(row => string.Equals(row.Section, section, StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (sectionRows.Count == 0)
            {
                continue;
            }

            worksheet.Cell(currentRow, 1).Value = section;
            worksheet.Cell(currentRow, 1).Style.Font.Bold = true;
            currentRow++;

            var headers = CollectHeaders(sectionRows);
            for (var columnIndex = 0; columnIndex < headers.Count; columnIndex++)
            {
                worksheet.Cell(currentRow, columnIndex + 1).Value = headers[columnIndex];
                worksheet.Cell(currentRow, columnIndex + 1).Style.Font.Bold = true;
            }

            worksheet.Cell(currentRow, headers.Count + 1).Value = "Quantity";
            worksheet.Cell(currentRow, headers.Count + 1).Style.Font.Bold = true;
            currentRow++;

            foreach (var row in sectionRows)
            {
                for (var columnIndex = 0; columnIndex < headers.Count; columnIndex++)
                {
                    worksheet.Cell(currentRow, columnIndex + 1).Value =
                        row.Values.TryGetValue(headers[columnIndex], out var value) ? value : string.Empty;
                }

                worksheet.Cell(currentRow, headers.Count + 1).Value = row.Quantity;
                currentRow++;
            }

            currentRow++;
        }

        worksheet.SheetView.FreezeRows(1);
        worksheet.Columns().AdjustToContents();
        workbook.SaveAs(output);
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
}
