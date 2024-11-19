using Microsoft.AspNetCore.OpenApi;
using Microsoft.OpenApi.Models;

namespace Introl.Tools.Api.Authorization;

internal sealed class AuthHeaderTransformer
    : IOpenApiDocumentTransformer
{
    private const string ApiKeySchemeId = "ApiKey";

    public Task TransformAsync(OpenApiDocument document, OpenApiDocumentTransformerContext context,
        CancellationToken cancellationToken)
    {
        var requirements = new Dictionary<string, OpenApiSecurityScheme>
        {
            [ApiKeySchemeId] = new OpenApiSecurityScheme
            {
                Type = SecuritySchemeType.ApiKey,
                Scheme = AuthorizationConstants.ApiKeyHeader,
                In = ParameterLocation.Header,
                Name = AuthorizationConstants.ApiKeyHeader
            }
        };

        document.Components ??= new OpenApiComponents();
        document.Components.SecuritySchemes = requirements;

        foreach (var operation in document.Paths.Values.SelectMany(path => path.Operations))
        {
            operation.Value.Security.Add(new OpenApiSecurityRequirement
            {
                [new OpenApiSecurityScheme
                {
                    Reference = new OpenApiReference
                    {
                        Id = ApiKeySchemeId,
                        Type = ReferenceType.SecurityScheme
                    }
                }] = Array.Empty<string>()
            });
        }
        return Task.CompletedTask;
    }
}
