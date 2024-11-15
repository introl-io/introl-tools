using Introl.Timesheets.Api.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc.Testing;

namespace Introl.Timesheets.Api.Tests.Acceptance;

internal class AcceptanceTestsWebHost : WebApplicationFactory<Program>
{
    protected override void ConfigureWebHost(IWebHostBuilder builder)
    {
        Environment.SetEnvironmentVariable(AuthorizationConstants.ApiKeyEnvVariable, AcceptanceTestConstants.ApiKey);
    }
}
