namespace Introl.Timesheets.Console.models;

public class Employee
{
    public required string Name { get; init; }
    public required Dictionary<DayOfTheWeek, WorkDayHours> WorkDays { get; init; }
    public required decimal RegularHoursRate { get; init; }
    public required decimal OvertimeHoursRate { get; init; }
}