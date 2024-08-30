using System.Net;
using ClosedXML.Excel;
using Introl.Timesheets.Api.Authorization;
using Introl.Timesheets.Api.Tests.Acceptance.Utils;
using Xunit;

namespace Introl.Timesheets.Api.Tests.Acceptance.Controllers;

public class RackLabelControllerTests
{
    private readonly AcceptanceTestsWebHost _webHost = new();
    private readonly HttpClient _httpClient;

    public RackLabelControllerTests()
    {
        _httpClient = _webHost.CreateClient();
        _httpClient.DefaultRequestHeaders.Add(AuthorizationConstants.ApiKeyHeader, AcceptanceTestConstants.ApiKey);
    }

    [Fact]
    public async Task RackLabels_GivenKnownInput_GivesKnownOutput()
    {
        await using var inputFileStream = File.Open("./Resources/RackLabels/input.xlsx", FileMode.Open);

        var request =
            new MultipartFormDataContent { { new StreamContent(inputFileStream), "input", "input.xlsx" } };
        var response = await _httpClient.PostAsync("/api/rack-labels/create", request);

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        var contentDisposition = response.Content.Headers.ContentDisposition;
        Assert.Equal("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            response.Content.Headers.ContentType?.MediaType);
        Assert.Equal("PortLabels.xlsx", contentDisposition?.FileName);
        await using var responseStream = await response.Content.ReadAsStreamAsync();

        await using var expectedFileStream = File.Open("./Resources/RackLabels/expected_output.xlsx", FileMode.Open);
        var expectedWorkbook = new XLWorkbook(expectedFileStream);
        var responseWorkbook = new XLWorkbook(responseStream);

        AcceptanceTestUtils.CompareWorkbooks(responseWorkbook, expectedWorkbook);
    }

    [Fact]
    public async Task Team_WhenUploadUnsupportedFileTime_ReturnsBadRequest()
    {
        var request =
            new MultipartFormDataContent { { new StringContent(""), "input", "timesheet_input.pdf" } };
        var response = await _httpClient.PostAsync("/api/rack-labels/create", request);

        Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
        Assert.Equal("Unsupported file type: .pdf. Please upload a .xlsx file.", await response.Content.ReadAsStringAsync());
    }

}
