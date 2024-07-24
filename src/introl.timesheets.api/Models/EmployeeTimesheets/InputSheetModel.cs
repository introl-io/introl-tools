using ClosedXML.Excel;

namespace Introl.Timesheets.Api.Models.EmployeeTimesheets;

public class InputSheetModel
{
    public required DateOnly StartDate { get; init; }
    public required DateOnly EndDate { get; init; }
    public required IList<Employee> Employees { get; init; }
    public required IXLWorksheet RawTimesheetsWorksheet { get; init; }
}
