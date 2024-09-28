using System.Net;
using System.Net.Http.Json;
using ClosedXML.Excel;
using Introl.Timesheets.Api.Authorization;
using Introl.Timesheets.Api.Models;
using Introl.Timesheets.Api.Tests.Acceptance.Utils;
using Microsoft.AspNetCore.Http;
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

        using var content = new MultipartFormDataContent();
        content.Add(new StreamContent(inputFileStream), "File", "input.xlsx");
        content.Add(new StringContent("{F}.R{G}.{I}.{J}"), "SourcePortLabelFormat");
        content.Add(new StringContent("{R}.R{S}.{U}.{V}"), "DestinationPortLabelFormat");
        using var request = new HttpRequestMessage(HttpMethod.Post, "/api/rack-labels/create");
        request.Content = content;
        var response = await _httpClient.SendAsync(request);

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        var contentDisposition = response.Content.Headers.ContentDisposition;
        Assert.Equal("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            response.Content.Headers.ContentType?.MediaType);
        Assert.Equal("PortLabels.xlsx", contentDisposition?.FileName);
        await using var responseStream = await response.Content.ReadAsStreamAsync();

        // await using var fileStream = new FileStream("./Resources/RackLabels/output.xlsx", FileMode.Create, FileAccess.Write, FileShare.None);
        // await responseStream.CopyToAsync(fileStream);

        await using var expectedFileStream = File.Open("./Resources/RackLabels/expected_output.xlsx", FileMode.Open);
        var expectedWorkbook = new XLWorkbook(expectedFileStream);
        var responseWorkbook = new XLWorkbook(responseStream);

        AcceptanceTestUtils.CompareWorkbooks(responseWorkbook, expectedWorkbook);
    }

    [Fact]
    public async Task Team_WhenUploadUnsupportedFileTime_ReturnsBadRequest()
    {
        await using var inputFileStream = File.Open("./Resources/RackLabels/input.xlsx", FileMode.Open);

        using var content = new MultipartFormDataContent();
        content.Add(new StreamContent(inputFileStream), "File", "input.pdf");
        content.Add(new StringContent("{F}.R{G}.{I}.{J}"), "SourcePortLabelFormat");
        content.Add(new StringContent("{R}.R{S}.{U}.{V}"), "DestinationPortLabelFormat");
        using var request = new HttpRequestMessage(HttpMethod.Post, "/api/rack-labels/create");
        request.Content = content;
        var response = await _httpClient.SendAsync(request);

        Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
        var x = await response.Content.ReadAsStringAsync();
        Assert.Equal("Unsupported file type: .pdf. Please upload a .xlsx file.", await response.Content.ReadAsStringAsync());
    }

}
