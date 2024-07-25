namespace Introl.Timesheets.Api.Models.ActivityCodeTimesheets;

public record ActivityCodeHours
{
    public required DateTime StartTime { get; init; }
    public required double Hours { get; init; }
    public required string ActivityCode { get; init; }
}
