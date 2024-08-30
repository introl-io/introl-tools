using System.Net;
using ClosedXML.Excel;
using FluentAssertions;
using Introl.Timesheets.Api.Authorization;
using Introl.Timesheets.Api.Tests.Acceptance.Utils;
using Introl.Tools.Common.Utils;
using Xunit;

namespace Introl.Timesheets.Api.Tests.Acceptance.Controllers;

public class TimesheetControllerTests
{
    private readonly AcceptanceTestsWebHost _webHost = new();
    private readonly HttpClient _httpClient;

    public TimesheetControllerTests()
    {
        _httpClient = _webHost.CreateClient();
        _httpClient.DefaultRequestHeaders.Add(AuthorizationConstants.ApiKeyHeader, AcceptanceTestConstants.ApiKey);
    }

    [Fact]
    public async Task Team_GivenKnownInput_GivesKnownOutput()
    {
        await using var inputFileStream = File.Open("./Resources/Timesheets/Team/Success/timesheet_input.xlsx", FileMode.Open);

        var request =
            new MultipartFormDataContent { { new StreamContent(inputFileStream), "input", "timesheet_input.xlsx" } };
        var response = await _httpClient.PostAsync("/api/timesheet/team", request);

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        var expectedFileName = "\"Weekly Timesheet - Introl.io 2024.07.08 - 2024.07.14.xlsx\"";
        var contentDisposition = response.Content.Headers.ContentDisposition;
        Assert.Equal("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            response.Content.Headers.ContentType?.MediaType);
        Assert.Equal(expectedFileName, contentDisposition?.FileName);
        await using var responseStream = await response.Content.ReadAsStreamAsync();

        await using var expectedFileStream = File.Open("./Resources/Timesheets/Team/Success/expected_output.xlsx", FileMode.Open);
        var expectedWorkbook = new XLWorkbook(expectedFileStream);
        var responseWorkbook = new XLWorkbook(responseStream);

        AcceptanceTestUtils.CompareWorkbooks(responseWorkbook, expectedWorkbook);
    }

    [Fact]
    public async Task Team_WhenUploadUnsupportedFileTime_ReturnsBadRequest()
    {
        var request =
            new MultipartFormDataContent { { new StringContent(""), "input", "timesheet_input.pdf" } };
        var response = await _httpClient.PostAsync("/api/timesheet/team", request);

        Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
        Assert.Equal("Unsupported file type: .pdf. Please upload a .xlsx file.", await response.Content.ReadAsStringAsync());
    }

    [Fact]
    public async Task ActivityCode_GivenKnownInput_GivesKnownOutput()
    {
        await using var inputFileStream = File.Open("./Resources/Timesheets/ActivityCode/Success/timesheet_input.xlsx", FileMode.Open);

        var request =
            new MultipartFormDataContent { { new StreamContent(inputFileStream), "input", "timesheet_input.xlsx" } };
        var response = await _httpClient.PostAsync("/api/timesheet/activity-code", request);

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);

        var expectedFileName = "\"MEM-Q1368 Timesheet - Introl.io 2024.07.15 - 2024.07.21.xlsx\"";
        var contentDisposition = response.Content.Headers.ContentDisposition;
        Assert.Equal("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            response.Content.Headers.ContentType?.MediaType);
        Assert.Equal(expectedFileName, contentDisposition?.FileName);
        await using var responseStream = await response.Content.ReadAsStreamAsync();

        await using var expectedFileStream = File.Open("./Resources/Timesheets/ActivityCode/Success/expected_output.xlsx", FileMode.Open);
        var expectedWorkbook = new XLWorkbook(expectedFileStream);
        var responseWorkbook = new XLWorkbook(responseStream);

        AcceptanceTestUtils.CompareWorkbooks(responseWorkbook, expectedWorkbook);
    }

    [Fact]
    public async Task ActivityCode_WhenUploadUnsupportedFileTime_ReturnsBadRequest()
    {
        var request =
            new MultipartFormDataContent { { new StringContent(""), "input", "timesheet_input.pdf" } };
        var response = await _httpClient.PostAsync("/api/timesheet/activity-code", request);

        Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
        Assert.Equal("Unsupported file type: .pdf. Please upload a .xlsx file.", await response.Content.ReadAsStringAsync());
    }


}
