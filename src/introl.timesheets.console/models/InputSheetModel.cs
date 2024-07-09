using ClosedXML.Excel;
using Introl.Timesheets.Console.constants;

namespace Introl.Timesheets.Console.models;

public class InputSheetModel
{
    public required DateOnly StartDate { get; init; }
    public required DateOnly EndDate { get; init; }
    public required IEnumerable<Employee> Employees { get; init; }
    public required IXLWorksheet RawTimesheetsWorksheet { get; init; }

    public int WeekNumber
    {
        get
        {
            var days = StartDate.DayNumber - DateConstants.ProjectStartDate.DayNumber;
            return (days / 7) + 1;
        }
    }
}
