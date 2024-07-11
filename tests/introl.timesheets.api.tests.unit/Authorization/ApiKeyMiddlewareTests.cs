using Introl.Timesheets.Api.Authorization;
using Microsoft.AspNetCore.Http;
using Xunit;

namespace Introl.Timesheets.Api.Tests.Unit.Authorization;

public class ApiKeyMiddlewareTests
{
    private const string ApiKey = "test-api-key";
    
    private readonly ApiKeyMiddleware _sut = new();

    public ApiKeyMiddlewareTests()
    {
        Environment.SetEnvironmentVariable(AuthorizationConstants.ApiKeyEnvVariable, ApiKey);
    }
    
    [Theory]
    [InlineData("/swagger", "not-api-key", true)]
    [InlineData("/swagger/index.html", null, true)]
    [InlineData("/other-path", "not-api-key", false)]
    [InlineData("/other-path", null, false)]
    [InlineData("/other-path", ApiKey, true)]
    public async Task InvokeAsync_ProcessesRequestsCorrectly(string path, string? apiKey, bool valid)
    {
        var ctx = new DefaultHttpContext();
        ctx.Request.Path = path;
        ctx.Request.Headers[AuthorizationConstants.ApiKeyHeader] = apiKey;
        RequestDelegate next = (HttpContext hc) =>
        {
            hc.Response.StatusCode = StatusCodes.Status200OK;
            return Task.CompletedTask;
        };

        await _sut.InvokeAsync(ctx, next);

        ctx.Response.StatusCode = valid ? StatusCodes.Status200OK : StatusCodes.Status401Unauthorized;
    }
}
