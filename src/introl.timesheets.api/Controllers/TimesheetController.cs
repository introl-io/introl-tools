using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Introl.Timesheets.Api.services;
using Microsoft.AspNetCore.Mvc;

namespace Introl.Timesheets.Api.Controllers;

[Route("api/timesheet")]
[ApiController]
public class TimesheetController(
    IWorksheetReader worksheetReader, 
    IWorksheetWriter worksheetWriter) : Controller
{
    [HttpPost("process")]
    public IActionResult ProcessTimeSheet(IFormFile model)
    {
        Console.WriteLine("Processing timesheet...");
        using var x = model.OpenReadStream();
        var workbook = new XLWorkbook(model.OpenReadStream());
        
        Console.WriteLine("Starting reader");
        var inputSheetModel = worksheetReader.Process(workbook);
        Console.WriteLine("After reader");
        using var output = new XLWorkbook();
        worksheetWriter.Process(inputSheetModel, output);
        Console.WriteLine("After writer");
        
        using var stream = new MemoryStream();
        output.SaveAs(stream);
        Console.WriteLine("After save");
        return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"{model.FileName}_processed.xlsx");
    }
}
