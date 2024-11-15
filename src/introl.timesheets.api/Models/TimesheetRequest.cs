using Swashbuckle.AspNetCore.Annotations;

namespace Introl.Timesheets.Api.Models;

public record TimesheetRequest
{
    public required IFormFile File { get; init; }
    public required bool CalculateOvertime { get; init; }
}
