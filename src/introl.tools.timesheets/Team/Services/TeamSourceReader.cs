using ClosedXML.Excel;
using Introl.Tools.Common.Extensions;
using Introl.Tools.Timesheets.Team.Models;

namespace Introl.Tools.Timesheets.Team.Services;

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
            Employees = GetEmployees(teamSummarySheet, startDate, endDate),
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

    private IList<TeamEmployee> GetEmployees(IXLWorksheet worksheet, DateOnly startDate, DateOnly endDate)
    {
        var employeeRow = GetFirstEmployeeRow(worksheet);
        var ratesCol = RatesColumn(worksheet);
        var employees = new List<TeamEmployee>();
        do
        {
            employees.Add(GetEmployee(worksheet, employeeRow, ratesCol, startDate, endDate, out var numRowsUsedByEmployee));
            employeeRow += numRowsUsedByEmployee;
        } while (!string.IsNullOrEmpty(worksheet.Cell(employeeRow, 1).GetString()));

        return employees;
    }

    private TeamEmployee GetEmployee(IXLWorksheet worksheet, int employeeRow, int ratesCol, DateOnly startDate, DateOnly endDate, out int numRowsUsedByEmployee)
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

        var workDays = new Dictionary<DateOnly, TeamEmployeeWorkDayHours>();
        var numDays = GetNumberOfDays(startDate, endDate);
        
        for(var i =0; i < numDays; i++)
        {
            var day = startDate.AddDays(i);
            workDays.Add(day, teamSourceParser.GetWorkdayHoursForEmployeeAndDay(worksheet, employeeRow, 3+i));
        }
        
        return new TeamEmployee
        {
            Name = name,
            RegularHoursRate = regularHoursRate,
            OvertimeHoursRate = overtimeHoursRate,
            WorkDays = workDays
        };
    }
    
    private int GetNumberOfDays(DateOnly startDate, DateOnly endDate)
    {
        return endDate.DayNumber - startDate.DayNumber + 1;
    }
}

public interface ITeamSourceReader
{
    TeamParsedSourceModel Process(XLWorkbook workbook);
}
