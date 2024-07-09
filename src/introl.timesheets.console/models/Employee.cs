namespace Introl.Timesheets.Console.models;

public class Employee
{
    public required string Name { get; init; }
    public required Dictionary<DayOfTheWeek, WorkDayHours> WorkDays { get; init; }
    public required decimal RegularHoursRate { get; init; }
    public required decimal OvertimeHoursRate { get; init; }
    public double TotalRegularHours => WorkDays.Sum(w => w.Value.RegularHours);
    public double TotalOvertimeHours => WorkDays.Sum(w => w.Value.OvertimeHours);
    public double TotalHours => TotalRegularHours + TotalOvertimeHours;
    public decimal TotalRegularBill => (decimal)TotalRegularHours * RegularHoursRate;
    public decimal TotalOvertimeBill => (decimal)TotalOvertimeHours * OvertimeHoursRate;
    public decimal TotalBill => TotalRegularBill + TotalOvertimeBill;
}