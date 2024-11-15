using ClosedXML.Excel;
using FluentAssertions;
using Introl.Tools.Common.Utils;
using Xunit;

namespace Introl.Timesheets.Api.Tests.Acceptance.Utils;

public static class AcceptanceTestUtils
{
    public static void CompareWorkbooks(XLWorkbook actual, XLWorkbook expected)
    {
        Assert.Equal(expected.Worksheets.Count(), actual.Worksheets.Count());

        for (var i = 1; i <= actual.Worksheets.Count(); i++)
        {
            var actualWorksheet = actual.Worksheet(i);
            var expectedWorksheet = expected.Worksheet(i);
            CompareWorksheets(actualWorksheet, expectedWorksheet, actualWorksheet.Name);
        }
    }

    private static void CompareWorksheets(IXLWorksheet actual, IXLWorksheet expected, string worksheetName)
    {
        actual.Rows().Count().Should().Be(expected.Rows().Count(), $"Row count mismatch in worksheet {worksheetName}");
        actual.Columns().Count().Should()
            .Be(expected.Columns().Count(), $"Row count mismatch in worksheet {worksheetName}");

        for (var i = 1; i <= actual.Rows().Count(); i++)
        {
            var actualRow = actual.Row(i);
            var expectedRow = expected.Row(i);
            CompareRows(actualRow, expectedRow, worksheetName, i);
        }
    }

    private static void CompareRows(IXLRow actual, IXLRow expected, string workSheetName, int rowNumber)
    {
        actual.CellsUsed().Count().Should().Be(expected.CellsUsed().Count(),
            $"Cell count mismatch in worksheet {workSheetName} row {rowNumber}");

        for (var i = 1; i <= actual.Cells().Count(); i++)
        {
            var actualCell = actual.Cell(i);
            var expectedCell = expected.Cell(i);
            actualCell.Value.ToString().Trim().Should().Be(expectedCell.Value.ToString().Trim(),
                $"Cell value mismatch in worksheet {workSheetName} cell {ExcelUtils.ToExcelColumn(i)}{rowNumber}");
        }
    }
}
