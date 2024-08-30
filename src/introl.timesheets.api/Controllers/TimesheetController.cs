using System.ComponentModel.DataAnnotations;
using Introl.Tools.Common.Enums;
using Introl.Tools.Timesheets.ActivityCode.Services;
using Introl.Tools.Timesheets.Team.Services;
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
                    ProcessingFailureReasons.UnsupportedFileType => BadRequest(
                        failReason.Message),
                    _ => throw new Exception("Unhandled fail reason getting results")
                };
            });
    }

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
                    ProcessingFailureReasons.UnsupportedFileType => BadRequest(
                        failReason.Message),
                    _ => throw new Exception("Unhandled fail reason getting results")
                };
            });
    }
}
