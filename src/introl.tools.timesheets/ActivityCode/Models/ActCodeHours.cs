namespace Introl.Tools.Timesheets.ActivityCode.Models;

public record ActCodeHours
{
    public required DateTime StartTime { get; init; }
    public required double Hours { get; init; }
    public required string ActivityCode { get; init; }
}
