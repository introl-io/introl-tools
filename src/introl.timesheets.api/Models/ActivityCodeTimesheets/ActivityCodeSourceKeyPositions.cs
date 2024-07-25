namespace Introl.Timesheets.Api.Models.ActivityCodeTimesheets;

public class ActivityCodeSourceKeyPositions
{
    public required int TitleRow { get; init; }
    public required int NameCol { get; init; }
    public required int MemberCodeCol { get; init; }
    public required int TrackedHoursCol { get; init; }
    public required int DateCol { get; init; }
    public required int StartTimeCol { get; init; }
    public required int ActivityCodeCol { get; init; }
    public required int BillableRateCol { get; init; }
    public required int TotalHoursRow { get; init; }
}
