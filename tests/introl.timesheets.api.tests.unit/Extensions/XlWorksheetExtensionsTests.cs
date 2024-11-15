using ClosedXML.Excel;
using Introl.Tools.Common.Extensions;
using Xunit;

namespace Introl.Timesheets.Api.Tests.Unit.Extensions;

public class XlWorksheetExtensionsTests
{
    [Fact]
    public void FindSingleCellByValue_WhenValueIsFound_ReturnsCell()
    {
        // Arrange
        var worksheet = new XLWorkbook().Worksheets.Add("Sheet1");
        worksheet.Cell("A1").Value = "Not me";
        worksheet.Cell("C3").Value = "find me";

        // Act
        var cell = worksheet.FindSingleCellByValue("FIND Me");

        // Assert
        Assert.Equal("find me", cell.GetString());
        Assert.Equal("C", cell.Address.ColumnLetter);
        Assert.Equal(3, cell.Address.RowNumber);
    }

    [Fact]
    public void FindSingleCellByValue_WhenValueIsNotFound_ThrowsException()
    {
        // Arrange
        var worksheet = new XLWorkbook().Worksheets.Add("Sheet1");
        worksheet.Cell("A1").Value = "Not me";
        worksheet.Cell("C3").Value = "Also Not me";

        // Act
        Assert.Throws<ArgumentNullException>(() => worksheet.FindSingleCellByValue("find me"));
    }

    [Fact]
    public void FindSingleCellByValue_WhenMultipleCellsFound_ThrowsException()
    {
        // Arrange
        var worksheet = new XLWorkbook().Worksheets.Add("Sheet1");
        worksheet.Cell("A1").Value = "find ME";
        worksheet.Cell("C3").Value = "FIND me";

        // Act
        Assert.Throws<InvalidOperationException>(() => worksheet.FindSingleCellByValue("find me"));
    }
}
