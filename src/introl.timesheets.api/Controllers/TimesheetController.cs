using System.ComponentModel.DataAnnotations;
using Introl.Timesheets.Api.Enums;
using Introl.Timesheets.Api.Timesheets.ActivityCode.Services;
using Introl.Timesheets.Api.Timesheets.Team.Services;
using Microsoft.AspNetCore.Mvc;

namespace Introl.Timesheets.Api.Controllers;

[Route("api/timesheet")]
[ApiController]
public class TimesheetController(
    IEmployeeEmployeeTimesheetProcessor employeeEmployeeEmployeeEmployeeTimesheetProcessor,
    IActCodeTimesheetProcessor actCodeTimesheetProcessor) : Controller
{
    [HttpPost("team")]
    public IActionResult Team([Required] IFormFile input)
    {
        var response = employeeEmployeeEmployeeEmployeeTimesheetProcessor.ProcessTimesheet(input);

        return response.Match<IActionResult>(
            results =>
            {
                using var stream = new MemoryStream();
                return File(results.WorkbookBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    results.Name);
            },
            failReason =>
            {
                return failReason.FailureReason switch
                {
                    TimesheetProcessingFailureReasons.UnsupportedFileType => BadRequest(
                        failReason.Message),
                    _ => throw new Exception("Unhandled fail reason getting results")
                };
            });
    }

    [ApiExplorerSettings(IgnoreApi = true)]
    [HttpPost("activity-code")]
    public IActionResult ActivityCode([Required] IFormFile input)
    {
        var response = actCodeTimesheetProcessor.ProcessTimesheet(input);

        return response.Match<IActionResult>(
            results =>
            {
                using var stream = new MemoryStream();
                return File(results.WorkbookBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    results.Name);
            },
            failReason =>
            {
                return failReason.FailureReason switch
                {
                    TimesheetProcessingFailureReasons.UnsupportedFileType => BadRequest(
                        failReason.Message),
                    _ => throw new Exception("Unhandled fail reason getting results")
                };
            });
    }
}
