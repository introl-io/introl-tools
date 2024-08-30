using ClosedXML.Excel;
using FluentAssertions;
using Introl.Tools.Timesheets.Team.Services;
using Xunit;

namespace Introl.Timesheets.Api.Tests.Unit.Timesheets.Team;

public class TeamSourceParserTests
{
    private TeamSourceParser sut = new();

    [Fact]
    public void GetStartAndEndDate_MissingWeekCell_ThrowsException()
    {
        // Arrange
        var worksheet = new XLWorkbook().Worksheets.Add("Sheet1");

        Assert.Throws<ArgumentNullException>(() => sut.GetStartAndEndDate(worksheet));
    }

    [Fact]
    public void GetStartAndEndDate_ParsesDateCorrectly()
    {
        // Arrange
        var worksheet = new XLWorkbook().Worksheets.Add("Sheet1");
        worksheet.Cell("A1").Value = "week";
        worksheet.Cell("B1").Value = "01 July 2024 - 07 July 2024";

        var expectedStartDate = new DateOnly(2024, 7, 1);
        var expectedEndDate = new DateOnly(2024, 7, 7);

        var (startDate, endDate) = sut.GetStartAndEndDate(worksheet);
        startDate.Should().Be(expectedStartDate);
        endDate.Should().Be(expectedEndDate);
    }

    [Theory]
    [InlineData("12:00", "00:00", 12, 0)]
    [InlineData("12:19", "00:20", 12, 0.5)]
    [InlineData("12:39", "00:40", 12.5, 1)]
    [InlineData("12", "-", 12, 0)]
    public void GetWorkdayHoursForEmployeeAndDay_ReturnsCorrectHours(string regularHours, string overtimeHours, double expectedRegularHours, double expectedOvertimeHours)
    {
        // Arrange
        var worksheet = new XLWorkbook().Worksheets.Add("Sheet1");
        worksheet.Cell("A1").Value = "Type";
        worksheet.Cell("A2").Value = "Regular Hours";
        worksheet.Cell("A3").Value = "Weekly OT";
        worksheet.Cell("B2").Value = regularHours;
        worksheet.Cell("B3").Value = overtimeHours;

        // Act
        var result = sut.GetWorkdayHoursForEmployeeAndDay(worksheet, 1, 2);

        // Assert
        result.RegularHours.Should().Be(expectedRegularHours);
        result.OvertimeHours.Should().Be(expectedOvertimeHours);
    }

    [Theory]
    [InlineData("53.32", "12.00", 53.32, 12)]
    [InlineData("", " ", 0, 0)]
    public void GetEmployeeRates_ReturnsCorrectRates(string regularRate, string overtimeRate, decimal expectedRegularRate, decimal expectedOvertimeRate)
    {
        // Arrange
        var worksheet = new XLWorkbook().Worksheets.Add("Sheet1");
        worksheet.Cell("A1").Value = "Type";
        worksheet.Cell("A2").Value = "Regular Hours";
        worksheet.Cell("A3").Value = "Weekly OT";
        worksheet.Cell("B2").Value = regularRate;
        worksheet.Cell("B3").Value = overtimeRate;

        // Act
        var result = sut.GetEmployeeRates(worksheet, 1, 2);

        // Assert
        result.regularHoursRate.Should().Be(expectedRegularRate);
        result.overtimeRate.Should().Be(expectedOvertimeRate);
    }
}
