using ClosedXML.Excel;

namespace Introl.Timesheets.Api.Models;

public class CellToAdd
{
    public required int Column { get; init; }
    public required int Row { get; init; }
    public required string Value { get; init; }
    public CellValueType ValueType = CellValueType.Value;
    public bool Bold = false;
    public XLColor Color = XLColor.White;
    
    public string? NumberFormat = null;

    public enum CellValueType
    {
        Value,
        Formula
    }
}
