using System.ComponentModel.DataAnnotations;
using Introl.Timesheets.Api.Models;
using Introl.Tools.Common.Enums;
using Introl.Tools.Timesheets.ActivityCode.Models;
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
    public IActionResult Team([FromForm] TimesheetRequest request)
    {
        var response = employeeEmployeeEmployeeEmployeeTimesheetProcessor.ProcessTimesheet(request.File);

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
    public IActionResult ActivityCode([FromForm] TimesheetRequest request)
    {
        var processRequest = new ActCodeProcessRequest
        {
            File = request.File,
            CalculateOvertime = request.CalculateOvertime
        };
        var response = actCodeTimesheetProcessor.ProcessTimesheet(processRequest);

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
