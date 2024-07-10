using Introl.Timesheets.Api;
using Introl.Timesheets.Api.services;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
builder.Services.AddScoped<IWorksheetReader, WorksheetReader>();
builder.Services.AddScoped<IWorksheetReaderHelper, WorksheetReaderReaderHelper>();
builder.Services.AddScoped<IWorksheetWriterHelper, WorksheetWriterHelper>();
builder.Services.AddScoped<IWorksheetWriter, WorksheetWriter>();
builder.Services.AddLogging();
var app = builder.Build();


app.UseSwagger();
app.UseSwaggerUI();


app.MapControllers();
app.UseHttpsRedirection();


app.Run();
