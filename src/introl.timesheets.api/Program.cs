using System.Collections;
using Introl.Timesheets.Api.Authorization;
using Introl.Timesheets.Api.Constants;
using Introl.Timesheets.Api.Services;
using Microsoft.OpenApi.Models;

var builder = WebApplication.CreateBuilder(args);
builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(opts =>
{
    opts.AddSecurityDefinition("ApiKey",
        new OpenApiSecurityScheme
        {
            Description = "The API Key to access the API.",
            Type = SecuritySchemeType.ApiKey,
            Name = AuthorizationConstants.ApiKeyHeader,
            In = ParameterLocation.Header,
            Scheme = "ApiKeyScheme"
        });

    var apiKeyScheme = new OpenApiSecurityScheme
    {
        Reference = new OpenApiReference { Type = ReferenceType.SecurityScheme, Id = "ApiKey" }
    };

    opts.AddSecurityRequirement(new OpenApiSecurityRequirement { [apiKeyScheme] = new List<string>() });


    opts.AddSecurityRequirement(new OpenApiSecurityRequirement { [apiKeyScheme] = new List<string>() });

});
builder.Services.AddScoped<IWorksheetReader, WorksheetReader>();
builder.Services.AddScoped<IWorksheetReaderHelper, WorksheetReaderHelper>();
builder.Services.AddScoped<IWorksheetWriterHelper, WorksheetWriterHelper>();
builder.Services.AddScoped<IWorksheetWriter, WorksheetWriter>();
builder.Services.AddScoped<ApiKeyMiddleware>();
builder.Services.AddLogging();

var app = builder.Build();
app.UseMiddleware<ApiKeyMiddleware>();
app.UseSwagger();
app.UseSwaggerUI();

app.MapControllers();
app.UseHttpsRedirection();


app.Run();
