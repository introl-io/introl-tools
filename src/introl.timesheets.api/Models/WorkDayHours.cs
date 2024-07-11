namespace Introl.Timesheets.Api.Models;

public class WorkDayHours
{
    public double RegularHours { get; init; }
    public double OvertimeHours { get; init; }
    public double TotalHours => RegularHours + OvertimeHours;
}