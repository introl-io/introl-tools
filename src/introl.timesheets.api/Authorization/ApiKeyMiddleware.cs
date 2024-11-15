namespace Introl.Timesheets.Api.Authorization;

public class ApiKeyMiddleware : IMiddleware
{

    public async Task InvokeAsync(HttpContext context, RequestDelegate next)
    {
        if (context.Request.Path.StartsWithSegments("/swagger"))
        {
            await next(context);
            return;
        }

        if (!context.Request.Headers.TryGetValue(AuthorizationConstants.ApiKeyHeader, out
                var requestApiKey))
        {
            context.Response.StatusCode = StatusCodes.Status401Unauthorized;
            return;
        }

        var apiKey = Environment.GetEnvironmentVariable(AuthorizationConstants.ApiKeyEnvVariable);
        if (requestApiKey != apiKey)
        {
            context.Response.StatusCode = StatusCodes.Status401Unauthorized;
            return;
        }

        await next(context);
    }
}
