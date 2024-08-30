using ClosedXML.Excel;
using Introl.Tools.Common.Models;

namespace Introl.Tools.Common.Extensions;

public static class XlWorksheetExtensions
{
    public static IXLCell FindSingleCellByValue(this IXLWorksheet worksheet, string value,
        bool ifMultipleTakeFirst = false)
    {
        var matchingCells = worksheet.CellsUsed(c => c.GetString().ToUpper() == value.ToUpper());
        if (!matchingCells.Any())
        {
            throw new ArgumentNullException($"No cell found with the value {value}");
        }

        if (matchingCells.Count() > 1)
        {
            if (ifMultipleTakeFirst)
            {
                return matchingCells.First();
            }

            throw new InvalidOperationException($"Multiple cells found with the value {value}");
        }

        return matchingCells.First();
    }

    public static void WriteCells(this IXLWorksheet worksheet, IEnumerable<CellToAdd> cells)
    {
        foreach (var cell in cells)
        {
            if (cell.Value.HasValue)
            {
                if (cell.ValueType == CellToAdd.CellValueType.Formula)
                {
                    worksheet.Cell(cell.Row, cell.Column).FormulaA1 = cell.Value.Value.ToString();
                }
                else
                {
                    worksheet.Cell(cell.Row, cell.Column).Value = cell.Value.Value;
                }
            }

            worksheet.Cell(cell.Row, cell.Column).Style.NumberFormat.Format = cell.NumberFormat;
            worksheet.Cell(cell.Row, cell.Column).Style.Font.Bold = cell.Bold;
            worksheet.Cell(cell.Row, cell.Column).Style.Alignment.WrapText = cell.Wrap;
            if (cell.Color is not null)
            {
                worksheet.Cell(cell.Row, cell.Column).Style.Fill.BackgroundColor = cell.Color;
            }

            worksheet.Cell(cell.Row, cell.Column).Style.Font.FontSize = cell.FontSize;
        }
    }
}
