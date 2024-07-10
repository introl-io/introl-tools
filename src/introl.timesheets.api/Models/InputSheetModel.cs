using ClosedXML.Excel;
using Introl.Timesheets.Api.Constants;

namespace Introl.Timesheets.Api.models;

public class InputSheetModel
{
    public required DateOnly StartDate { get; init; }
    public required DateOnly EndDate { get; init; }
    public required IEnumerable<Employee> Employees { get; init; }
    public required IXLWorksheet RawTimesheetsWorksheet { get; init; }
}
