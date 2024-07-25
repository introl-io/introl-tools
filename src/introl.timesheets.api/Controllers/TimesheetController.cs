using System.ComponentModel.DataAnnotations;
using Introl.Timesheets.Api.Enums;
using Introl.Timesheets.Api.Services;
using Introl.Timesheets.Api.Services.ActivityCodeTimesheets;
using Introl.Timesheets.Api.Services.EmployeeTimesheets;
using Microsoft.AspNetCore.Mvc;

namespace Introl.Timesheets.Api.Controllers;

[Route("api/timesheet")]
[ApiController]
public class TimesheetController(
    IEmployeeTimesheetProcessor employeeEmployeeTimesheetProcessor,
    IActivityCodeTimesheetProcessor activityCodeTimesheetProcessor) : Controller
{
    [HttpPost("employee")]
    public IActionResult Employee([Required] IFormFile input)
    {
        var response = employeeEmployeeTimesheetProcessor.ProcessTimesheet(input);

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
        var response = activityCodeTimesheetProcessor.ProcessTimesheet(input);

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
