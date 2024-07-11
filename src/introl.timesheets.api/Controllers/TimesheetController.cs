using ClosedXML.Excel;
using Introl.Timesheets.Api.Services;
using Microsoft.AspNetCore.Mvc;

namespace Introl.Timesheets.Api.Controllers;

[Route("api/timesheet")]
[ApiController]
public class TimesheetController(
    IWorksheetReader worksheetReader,
    IWorksheetWriter worksheetWriter) : Controller
{
    [HttpPost("process")]
    public IActionResult Process(IFormFile model)
    {
        var extension = Path.GetExtension(model.FileName);
        if (extension != ".xlsx")
        {
            return BadRequest($"Unsupported file type: {extension}");
        }
        var workbook = new XLWorkbook(model.OpenReadStream());

        var inputSheetModel = worksheetReader.Process(workbook);
        using var output = new XLWorkbook();
        worksheetWriter.Process(inputSheetModel, output);

        var dateFormat = "yyyy.MM.dd";
        var fileName =
            $"Weekly Timesheet - Introl.io {inputSheetModel.StartDate.ToString(dateFormat)} - {inputSheetModel.EndDate.ToString(dateFormat)}.xlsx";
        using var stream = new MemoryStream();
        output.SaveAs(stream);


        return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
    }
}
