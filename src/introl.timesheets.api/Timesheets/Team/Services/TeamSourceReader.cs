using ClosedXML.Excel;
using Introl.Timesheets.Api.Enums;
using Introl.Timesheets.Api.Extensions;
using Introl.Timesheets.Api.Timesheets.Team.Models;

namespace Introl.Timesheets.Api.Timesheets.Team.Services;

public class TeamSourceReader(ITeamSourceParser teamSourceParser) : ITeamSourceReader
{
    public TeamParsedSourceModel Process(XLWorkbook workbook)
    {
        var teamSummarySheet = workbook.Worksheets.Worksheet("Team Summary");
        var rawTimesheetsWorkSheet = workbook.Worksheets.Worksheet("Raw Timesheets");
        ArgumentNullException.ThrowIfNull(teamSummarySheet);
        var (startDate, endDate) = teamSourceParser.GetStartAndEndDate(teamSummarySheet);

        return new TeamParsedSourceModel
        {
            StartDate = startDate,
            EndDate = endDate,
            Employees = GetEmployees(teamSummarySheet),
            RawTimesheetsWorksheet = rawTimesheetsWorkSheet
        };
    }

    private int GetFirstEmployeeRow(IXLWorksheet worksheet)
    {
        var cell = worksheet.FindSingleCellByValue("name");
        return cell.Address.RowNumber + 1;
    }

    private int RatesColumn(IXLWorksheet worksheet)
    {
        var cell = worksheet.FindSingleCellByValue("RATES");
        return cell.Address.ColumnNumber;
    }

    private IList<TeamEmployee> GetEmployees(IXLWorksheet worksheet)
    {
        var employeeRow = GetFirstEmployeeRow(worksheet);
        var dayColDict = teamSourceParser.GetDayOfTheWeekColumnDictionary(worksheet);
        var ratesCol = RatesColumn(worksheet);
        var employees = new List<TeamEmployee>();
        do
        {
            employees.Add(GetEmployee(worksheet, employeeRow, dayColDict, ratesCol, out var numRowsUsedByEmployee));
            employeeRow += numRowsUsedByEmployee;
        } while (!string.IsNullOrEmpty(worksheet.Cell(employeeRow, 1).GetString()));

        return employees;
    }

    private TeamEmployee GetEmployee(IXLWorksheet worksheet, int employeeRow, IDictionary<DayOfTheWeek, int> dayDictionary, int ratesCol, out int numRowsUsedByEmployee)
    {
        var name = worksheet.Cell(employeeRow, 1).GetString();

        var (hasDoneRegularHours, hasDoneOvertimeHours) = teamSourceParser.GetTypesOfHoursEmployeeHasDone(worksheet, employeeRow);
        var (regularHoursRate, overtimeHoursRate) = teamSourceParser.GetEmployeeRates(worksheet, employeeRow, ratesCol);

        numRowsUsedByEmployee = 1;
        if (hasDoneRegularHours && hasDoneOvertimeHours)
        {
            numRowsUsedByEmployee += 2;
        }
        else if (hasDoneRegularHours || hasDoneOvertimeHours)
        {
            numRowsUsedByEmployee += 1;
        }

        var workDays = new Dictionary<DayOfTheWeek, TeamEmployeeWorkDayHours>();
        foreach (var day in Enum.GetValues(typeof(DayOfTheWeek)).Cast<DayOfTheWeek>())
        {
            workDays.Add(day, teamSourceParser.GetWorkdayHoursForEmployeeAndDay(worksheet, employeeRow, dayDictionary[day]));
        }

        return new TeamEmployee
        {
            Name = name,
            RegularHoursRate = regularHoursRate,
            OvertimeHoursRate = overtimeHoursRate,
            WorkDays = workDays
        };
    }
}

public interface ITeamSourceReader
{
    TeamParsedSourceModel Process(XLWorkbook workbook);
}
