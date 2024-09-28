using System.ComponentModel.DataAnnotations;
using Introl.Timesheets.Api.Models;
using Introl.Tools.Common.Enums;
using Introl.Tools.Racks.Models;
using Introl.Tools.Racks.Services;
using Microsoft.AspNetCore.Mvc;

namespace Introl.Timesheets.Api.Controllers;

[Route("api/rack-labels")]
[ApiController]
public class RackLabelsController(IRackProcessor rackProcessor) : Controller
{
    [HttpPost("create")]
    public IActionResult CreateLabels([FromForm] CreateLabelsRequest request)
    {
        var response = rackProcessor.Process(new ProcessFileRequest
        {
            File = request.File,
            SourcePortLabelFormat = request.SourcePortLabelFormat,
            DestinationPortLabelFormat = request.DestinationPortLabelFormat,
            HasHeadingRow = request.HasHeadingRow
        });
        return response.Match<IActionResult>(
            results =>
            {
                using var stream = new MemoryStream();
                return File(
                    results.WorkbookBytes,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
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
