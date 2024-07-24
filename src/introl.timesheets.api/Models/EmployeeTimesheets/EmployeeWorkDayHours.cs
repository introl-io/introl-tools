namespace Introl.Timesheets.Api.Models.EmployeeTimesheets;

public class EmployeeWorkDayHours
{
    public required double RegularHours { get; init; }
    public required double OvertimeHours { get; init; }
}
