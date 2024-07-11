using ClosedXML.Excel;

namespace Introl.Timesheets.Api.Constants;

public static class StyleConstants
{
    // Colours can be found here https://github.com/ClosedXML/ClosedXML/wiki/ClosedXML-Predefined-Colors
    public static XLColor Black => XLColor.SmokyBlack;
    public static XLColor LightGrey => XLColor.LightGray;
    public static XLColor DarkGrey => XLColor.DarkGray;
    public static int LargeFontSize => 14;
    public static string HourCellFormat => "0.00";
    public static string CurrencyWithSymbolCellFormat => $"$ {CurrencyCellFormat}";
    public static string CurrencyCellFormat => "#,##0.00";
}
