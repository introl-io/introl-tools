using Introl.Timesheets.Api;
using Introl.Timesheets.Api.Services;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
builder.Services.AddScoped<IWorksheetReader, WorksheetReader>();
builder.Services.AddScoped<IWorksheetReaderHelper, WorksheetReaderHelper>();
builder.Services.AddScoped<IWorksheetWriterHelper, WorksheetWriterHelper>();
builder.Services.AddScoped<IWorksheetWriter, WorksheetWriter>();
builder.Services.AddLogging();
var app = builder.Build();


app.UseSwagger();
app.UseSwaggerUI();


app.MapControllers();
app.UseHttpsRedirection();


app.Run();
