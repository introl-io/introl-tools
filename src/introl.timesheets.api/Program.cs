using Introl.Timesheets.Api.Authorization;
using Introl.Timesheets.Api.Services;
using Introl.Timesheets.Api.Services.ActivityCodeTimesheets;
using Introl.Timesheets.Api.Timesheets.Team.Services;
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

builder.Services.AddScoped<ITeamSourceReader, TeamSourceReader>();
builder.Services.AddScoped<ITeamSourceParser, TeamSourceParser>();
builder.Services.AddScoped<ITeamResultCellFactory, TeamResultCellFactory>();
builder.Services.AddScoped<ITeamResultWriter, TeamResultWriter>();
builder.Services.AddScoped<IEmployeeEmployeeTimesheetProcessor, EmployeeEmployeeEmployeeEmployeeTimesheetProcessor>();

builder.Services.AddScoped<IActivityCodeTimesheetProcessor, ActivityCodeTimesheetProcessor>();
builder.Services.AddScoped<IActivityCodeTimesheetReader, ActivityCodeTimesheetReader>();

builder.Services.AddScoped<ApiKeyMiddleware>();
builder.Services.AddLogging();

var app = builder.Build();
app.UseMiddleware<ApiKeyMiddleware>();
app.UseSwagger();
app.UseSwaggerUI();

app.MapControllers();
app.UseHttpsRedirection();


app.Run();
