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

    public static int ExcelColumnLetterToZeroBasedInt(string column)
    {
        column = column.ToUpper();
        int position = 0;

        foreach (var (letter, index) in column.Select((letter, index) => (letter, index)))
        {
            var positionMultiplier = column.Length - index - 1;
            var letterValue = letter - 'A';
            if (positionMultiplier > 0)
            {
                position += (letterValue + 1) * positionMultiplier * 26;
                break;
            }
            position += letterValue;
        }
        return position;
    }
}
