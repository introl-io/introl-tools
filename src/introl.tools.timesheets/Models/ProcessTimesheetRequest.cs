namespace Introl.Tools.Timesheets.Models;

public record ProcessTimesheetRequest
{
    public required IFormFile File { get; init; }
    public required bool CalculateOvertime { get; init; }
}
