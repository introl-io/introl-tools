namespace Introl.Tools.Timesheets.ActivityCode.Models;

public record ActCodeProcessRequest
{
    public required IFormFile File { get; init; }
    public required bool CalculateOvertime { get; init; }
}
