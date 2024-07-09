using ClosedXML.Excel;

namespace Introl.Timesheets.Console.services;

public class CellFinder : ICellFinder
{
    public IXLCell FindSingleCellByValue(IXLWorksheet worksheet, string value)
    {
        var matchingCells = worksheet.CellsUsed(c => c.GetString().ToUpper() == value.ToUpper());
        if(!matchingCells.Any())
        {
            throw new ArgumentNullException($"No cell found with the value {value}");
        }
        
        if(matchingCells.Count() > 1)
        {
            throw new InvalidOperationException($"Multiple cells found with the value {value}");
        }
        return matchingCells.First();
    }
}

public interface ICellFinder
{
    IXLCell FindSingleCellByValue(IXLWorksheet worksheet, string value);
}
