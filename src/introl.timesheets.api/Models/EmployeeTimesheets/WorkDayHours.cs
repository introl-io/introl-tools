namespace Introl.Timesheets.Api.Models.EmployeeTimesheets;

public class WorkDayHours
{
    public required double RegularHours { get; init; }
    public required double OvertimeHours { get; init; }
}
