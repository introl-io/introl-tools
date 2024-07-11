using System.Net;
using ClosedXML.Excel;
using FluentAssertions;
using Introl.Timesheets.Api.Authorization;
using Introl.Timesheets.Api.Constants;
using Xunit;

namespace Introl.Timesheets.Api.Tests.Acceptance.Controllers;

public class TimesheetControllerTests
{
    private readonly AcceptanceTestsWebHost _webHost = new();
    private HttpClient _httpClient;

    public TimesheetControllerTests()
    {
        _httpClient = _webHost.CreateClient();
        _httpClient.DefaultRequestHeaders.Add(AuthorizationConstants.ApiKeyHeader, AcceptanceTestConstants.ApiKey);
    }

    [Fact]
    public async Task Process_GivenKnownInput_GivesKnownOutput()
    {
        await using var inputFileStream = File.Open("./Resources/timesheet_input.xlsx", FileMode.Open);

        var request =
            new MultipartFormDataContent { { new StreamContent(inputFileStream), "model", "timesheet_input.xlsx" } };
        var response = await _httpClient.PostAsync("/api/timesheet/process", request);

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        var expectedFileName = "\"Weekly Timesheet - Introl.io 2024.07.08 - 2024.07.14.xlsx\"";
        var contentDisposition = response.Content.Headers.ContentDisposition;
        Assert.Equal("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            response.Content.Headers.ContentType?.MediaType);
        Assert.Equal(expectedFileName, contentDisposition?.FileName);
        await using var responseStream = await response.Content.ReadAsStreamAsync();

        await using var expectedFileStream = File.Open("./Resources/expected_output.xlsx", FileMode.Open);
        var expectedWorkbook = new XLWorkbook(expectedFileStream);
        var responseWorkbook = new XLWorkbook(responseStream);

        CompareWorkbooks(responseWorkbook, expectedWorkbook);
    }

    [Fact]
    public async Task Process_WhenUploadUnsupportedFileTime_ReturnsBadRequest()
    {
        var request =
            new MultipartFormDataContent { { new StringContent(""), "model", "timesheet_input.pdf" } };
        var response = await _httpClient.PostAsync("/api/timesheet/process", request);

        Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
        Assert.Equal($"Unsupported file type: .pdf", await response.Content.ReadAsStringAsync());
    }

    private void CompareWorkbooks(XLWorkbook actual, XLWorkbook expected)
    {
        Assert.Equal(actual.Worksheets.Count(), expected.Worksheets.Count());

        for (var i = 1; i <= actual.Worksheets.Count(); i++)
        {
            var actualWorksheet = actual.Worksheet(i);
            var expectedWorksheet = expected.Worksheet(i);
            CompareWorksheets(actualWorksheet, expectedWorksheet, actualWorksheet.Name);
        }
    }

    private void CompareWorksheets(IXLWorksheet actual, IXLWorksheet expected, string worksheetName)
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

    private void CompareRows(IXLRow actual, IXLRow expected, string workSheetName, int rowNumber)
    {
        actual.Cells().Count().Should().Be(expected.Cells().Count(),
            $"Cell count mismatch in worksheet {workSheetName} row {rowNumber}");

        for (var i = 1; i <= actual.Cells().Count(); i++)
        {
            var actualCell = actual.Cell(i);
            var expectedCell = expected.Cell(i);
            actualCell.Value.Should().Be(expectedCell.Value,
                $"Cell value mismatch in worksheet {workSheetName} row {rowNumber} cell {i}");
        }
    }
}
