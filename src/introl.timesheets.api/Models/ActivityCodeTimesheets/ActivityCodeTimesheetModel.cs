using ClosedXML.Excel;

namespace Introl.Timesheets.Api.Models.ActivityCodeTimesheets;

public class ActivityCodeTimesheetModel
{
    public required DateOnly StartDate { get; init; }
    public required DateOnly EndDate { get; init; }
    public required IDictionary<string, ActivityCodeEmployee> Employees { get; init; }
    public required IXLWorksheet InputWorksheet { get; init; }
}
