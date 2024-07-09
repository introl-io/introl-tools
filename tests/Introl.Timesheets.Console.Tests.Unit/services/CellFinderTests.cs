using ClosedXML.Excel;
using Introl.Timesheets.Console.services;
using Xunit;

namespace Introl.Timesheets.Console.Tests.Unit.services;

public class CellFinderTests
{
    private CellFinder sut;

    public CellFinderTests()
    {
        sut = new CellFinder();
    }
    
    [Fact]
    public void FindSingleCellByValue_WhenValueIsFound_ReturnsCell()
    {
        // Arrange
        var worksheet = new XLWorkbook().Worksheets.Add("Sheet1");
        worksheet.Cell("A1").Value = "Not me";
        worksheet.Cell("C3").Value = "find me";

        // Act
        var cell = sut.FindSingleCellByValue(worksheet, "FIND Me");

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
        Assert.Throws<ArgumentNullException>(() => sut.FindSingleCellByValue(worksheet, "find me"));
    }
    
    [Fact]
    public void FindSingleCellByValue_WhenMultipleCellsFound_ThrowsException()
    {
        // Arrange
        var worksheet = new XLWorkbook().Worksheets.Add("Sheet1");
        worksheet.Cell("A1").Value = "find ME";
        worksheet.Cell("C3").Value = "FIND me";

        // Act
        Assert.Throws<InvalidOperationException>(() => sut.FindSingleCellByValue(worksheet, "find me"));
    }
}
