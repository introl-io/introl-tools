using Introl.Tools.Api.Authorization;
using Introl.Tools.Racks.Services;
using Introl.Tools.Timesheets.ActivityCode.Services;
using Introl.Tools.Timesheets.Team.Services;

var builder = WebApplication.CreateBuilder(args);
builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddOpenApi(opts =>
{
    opts.AddDocumentTransformer<AuthHeaderTransformer>();
});

builder.Services.AddScoped<ITeamSourceReader, TeamSourceReader>();
builder.Services.AddScoped<ITeamSourceParser, TeamSourceParser>();
builder.Services.AddScoped<ITeamResultCellFactory, TeamResultCellFactory>();
builder.Services.AddScoped<ITeamResultWriter, TeamResultWriter>();
builder.Services.AddScoped<IEmployeeEmployeeTimesheetProcessor, EmployeeEmployeeEmployeeEmployeeTimesheetProcessor>();

builder.Services.AddScoped<IActCodeTimesheetProcessor, ActCodeTimesheetProcessor>();
builder.Services.AddScoped<IActCodeSourceReader, ActCodeSourceReader>();
builder.Services.AddScoped<IActCodeResultsWriter, ActCodeResultsWriter>();
builder.Services.AddScoped<IActCodeResultCellFactory, ActCodeResultCellFactory>();
builder.Services.AddScoped<IActCodeHoursProcessor, ActCodeHoursProcessor>();

builder.Services.AddScoped<IRackSourceReaderFactory, RackSourceReaderFactory>();
builder.Services.AddScoped<IRackSourceReader, RackSourceXlsxReader>();
builder.Services.AddScoped<IRackSourceReader, RackSourceCsvReader>();
builder.Services.AddScoped<IRackProcessor, RackProcessor>();
builder.Services.AddScoped<IRackCellFactory, RackCellFactory>();
builder.Services.AddScoped<IRackResultsWriter, RackResultsWriter>();

builder.Services.AddScoped<ApiKeyMiddleware>();
builder.Services.AddLogging();

var app = builder.Build();
app.UseMiddleware<ApiKeyMiddleware>();
const string swaggerSchemaFile = "/swagger/v1/openapi.json";
app.MapOpenApi(swaggerSchemaFile);

app.UseSwaggerUI(options =>
{
    options.SwaggerEndpoint(swaggerSchemaFile, "v1");
});

app.MapControllers();
app.UseHttpsRedirection();


app.Run();
