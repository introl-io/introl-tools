using ClosedXML.Excel;

namespace Introl.Timesheets.Api.Timesheets.Team.Models;

public class TeamParsedSourceModel
{
    public required DateOnly StartDate { get; init; }
    public required DateOnly EndDate { get; init; }
    public required IList<TeamEmployee> Employees { get; init; }
    public required IXLWorksheet RawTimesheetsWorksheet { get; init; }
}
