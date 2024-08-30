namespace Introl.Tools.Timesheets.Team.Models;

public class TeamEmployee
{
    public required string Name { get; init; }
    public required Dictionary<DateOnly, TeamEmployeeWorkDayHours> WorkDays { get; init; }
    public required decimal RegularHoursRate { get; init; }
    public required decimal OvertimeHoursRate { get; init; }
}
