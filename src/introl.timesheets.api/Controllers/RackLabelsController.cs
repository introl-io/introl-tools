using System.ComponentModel.DataAnnotations;
using Introl.Tools.Common.Enums;
using Introl.Tools.Racks.Services;
using Microsoft.AspNetCore.Mvc;

namespace Introl.Timesheets.Api.Controllers;

[Route("api/rack-labels")]
[ApiController]
public class RackLabelsController(IRackProcessor rackProcessor): Controller
{
    [HttpPost("create")]
    public IActionResult CreateLabels([Required] IFormFile input)
    {
        var response = rackProcessor.Process(input);
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
