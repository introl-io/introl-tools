using ClosedXML.Excel;

namespace Introl.Tools.Common.Models;

public class CellToAdd
{
    public required int Column { get; init; }
    public required int Row { get; init; }

    public XLCellValue? Value = null;
    public CellValueType ValueType = CellValueType.Value;
    public bool Bold = false;
    public bool Wrap = false;
    public XLColor? Color = null;
    public int FontSize = 11;

    public string? NumberFormat = null;

    public enum CellValueType
    {
        Value,
        Formula
    }
}
