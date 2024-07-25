namespace Introl.Timesheets.Api.Timesheets.Team.Models;

public class TeamEmployeeWorkDayHours
{
    public required double RegularHours { get; init; }
    public required double OvertimeHours { get; init; }
}
