using System.Net;
using System.Net.Mime;
using ClosedXML.Excel;
using FluentAssertions;
using Microsoft.AspNetCore.Mvc.Testing;
using Introl.Timesheets.Api;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Net.Http.Headers;
using Xunit;

namespace Introl.Timesheets.Api.Tests.Acceptance.Controllers;

public class TimesheetControllerTests
{
    private readonly WebApplicationFactory<Program> _webHost = new();
    private HttpClient _httpClient;

    public TimesheetControllerTests()
    {
        _httpClient = _webHost.CreateClient();
    }

    [Fact]
    public async Task ProcessTimesheet_GivenKnownInput_GivesKnownOutput()
    {
        var fileName = "timesheet_input.xslx";
        await using var inputFileStream = File.Open("./Resources/timesheet_input.xlsx", FileMode.Open);

        var response = await _httpClient.PostAsync("/api/timesheet/process",
            new MultipartFormDataContent { { new StreamContent(inputFileStream), "model", "timesheet_input.xlsx" } });

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        var contentDisposition = response.Content.Headers.ContentDisposition;
        Assert.Equal("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", response.Content.Headers.ContentType.MediaType);
        Assert.Equal("timesheet_input.xlsx_processed.xlsx", contentDisposition.FileName);
        await using var responseStream = await response.Content.ReadAsStreamAsync();

        await using var expectedFileStream = File.Open("./Resources/expected_output.xlsx", FileMode.Open);
        var expectedWorkbook = new XLWorkbook(expectedFileStream);
        var responseWorkbook = new XLWorkbook(responseStream);

        CompareWorkbooks(responseWorkbook, expectedWorkbook);
        // Assert.Equal(expectedOutputFileName, file.FileDownloadName);
    }

    private void CompareWorkbooks(XLWorkbook actual, XLWorkbook expected)
    {
        
        Assert.Equal(actual.Worksheets.Count(), expected.Worksheets.Count());
        
        for(var i = 1; i <= actual.Worksheets.Count(); i++)
        {
            var actualWorksheet = actual.Worksheet(i);
            var expectedWorksheet = expected.Worksheet(i);
            CompareWorksheets(actualWorksheet, expectedWorksheet, actualWorksheet.Name);
        }
    }
    
    private void CompareWorksheets(IXLWorksheet actual, IXLWorksheet expected, string worksheetName)
    {
        actual.Rows().Count().Should().Be(expected.Rows().Count(), $"Row count mismatch in worksheet {worksheetName}");
        actual.Columns().Count().Should().Be(expected.Columns().Count(), $"Row count mismatch in worksheet {worksheetName}");
        
        for(var i = 1; i <= actual.Rows().Count(); i++)
        {
            
            var actualRow = actual.Row(i);
            var expectedRow = expected.Row(i);
            CompareRows(actualRow, expectedRow, worksheetName, i);
        }
    }
    
    private void CompareRows(IXLRow actual, IXLRow expected, string workSheetName, int rowNumber)
    {
        actual.Cells().Count().Should().Be(expected.Cells().Count(), $"Cell count mismatch in worksheet {workSheetName} row {rowNumber}");
        
        for(var i = 1; i <= actual.Cells().Count(); i++)
        {
            var actualCell = actual.Cell(i);
            var expectedCell = expected.Cell(i);
            actualCell.Value.Should().Be(expectedCell.Value, $"Cell value mismatch in worksheet {workSheetName} row {rowNumber} cell {i}");
        }
    }
}
