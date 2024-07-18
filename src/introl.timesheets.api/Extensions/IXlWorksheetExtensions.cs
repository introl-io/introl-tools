using ClosedXML.Excel;

namespace Introl.Timesheets.Api.Extensions;

public static class XlWorksheetExtensions
{
    public static IXLCell FindSingleCellByValue(this IXLWorksheet worksheet, string value)
    {
        var matchingCells = worksheet.CellsUsed(c => c.GetString().ToUpper() == value.ToUpper());
        if (!matchingCells.Any())
        {
            throw new ArgumentNullException($"No cell found with the value {value}");
        }

        if (matchingCells.Count() > 1)
        {
            throw new InvalidOperationException($"Multiple cells found with the value {value}");
        }
        return matchingCells.First();
    }
}
