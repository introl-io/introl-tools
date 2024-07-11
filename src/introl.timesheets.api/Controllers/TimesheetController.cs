using ClosedXML.Excel;
using Introl.Timesheets.Api.Enums;
using Introl.Timesheets.Api.Models;
using Introl.Timesheets.Api.Services;
using Microsoft.AspNetCore.Mvc;
using OneOf;

namespace Introl.Timesheets.Api.Controllers;

[Route("api/timesheet")]
[ApiController]
public class TimesheetController(
    ITimesheetProcessor timesheetProcessor) : Controller
{
    [HttpPost("process")]
    public IActionResult Process(IFormFile model)
    {
        var response = timesheetProcessor.ProcessTimesheet(model);

        return response.Match<IActionResult>(
            results =>
            {
                using var stream = new MemoryStream();
                return File(results.WorkbookBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    results.Name);
            },
            failReason =>
            {
                return failReason switch
                {
                    TimesheetProcessingFailureReasons.UnsupportedFileType => BadRequest(
                        "Unsupported file type. Please upload a .xlsx file"),
                    _ => throw new Exception("Unhandled fail reason getting results")
                };
            });
    }
}
