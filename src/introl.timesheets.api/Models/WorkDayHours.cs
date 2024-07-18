namespace Introl.Timesheets.Api.Models;

public class WorkDayHours
{
    public required double RegularHours { get; init; }
    public required double OvertimeHours { get; init; }
}
