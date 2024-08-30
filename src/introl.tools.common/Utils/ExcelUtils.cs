namespace Introl.Tools.Common.Utils;

public static class ExcelUtils
{
    public static string ToExcelColumn(int columnNumber)
    {
        string columnName = "";

        while (columnNumber > 0)
        {
            int modulo = (columnNumber - 1) % 26;
            columnName = Convert.ToChar('A' + modulo) + columnName;
            columnNumber = (columnNumber - modulo) / 26;
        }

        return columnName;
    }

    public static string GetCellLocation(int columnNumber, int rowNumber)
    {
        return $"{ToExcelColumn(columnNumber)}{rowNumber}";
    }
}
