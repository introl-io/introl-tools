namespace Introl.Timesheets.Api.Models.ActivityCodeTimesheets;

public class ActivityCodeSourceColumns
{
    public required int Name { get; init; }
    public required int MemberCode { get; init; }
    public required int TrackedHours { get; init; }
    public required int Date { get; init; }
    public required int StartTime { get; init; }
    public required int ActivityCode { get; init; }
    public required int BillableRate { get; init; }
}
