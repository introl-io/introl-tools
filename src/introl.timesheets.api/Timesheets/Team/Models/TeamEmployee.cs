using Introl.Timesheets.Api.Enums;

namespace Introl.Timesheets.Api.Timesheets.Team.Models;

public class TeamEmployee
{
    public required string Name { get; init; }
    public required Dictionary<DayOfTheWeek, TeamEmployeeWorkDayHours> WorkDays { get; init; }
    public required decimal RegularHoursRate { get; init; }
    public required decimal OvertimeHoursRate { get; init; }
}
