using ClosedXML.Excel;
using Introl.Tools.Common.Extensions;
using Introl.Tools.Timesheets.Team.Constants;
using Introl.Tools.Timesheets.Team.Models;
using Introl.Tools.Timesheets.Utils;

namespace Introl.Tools.Timesheets.Team.Services;

public class TeamSourceParser : ITeamSourceParser
{
    public (DateOnly startDate, DateOnly endDate) GetStartAndEndDate(IXLWorksheet worksheet)
    {
        var weekCell = worksheet.FindSingleCellByValue(TeamSourceConstants.WeekCellValue);
        var dateString = weekCell.CellRight().GetString();
        var splitDates = dateString.Split(" - ");

        var startDate = DateOnly.Parse(splitDates[0]);
        var endDate = DateOnly.Parse(splitDates[1]);

        return (startDate, endDate);
    }

    public TeamEmployeeWorkDayHours GetWorkdayHoursForEmployeeAndDay(IXLWorksheet worksheet, int employeeRow, int dayColumn)
    {

        var (hasDoneRegularHours, hasDoneOvertimeHours) = GetTypesOfHoursEmployeeHasDone(worksheet, employeeRow);
        var overtimeIncrement = hasDoneRegularHours ? 2 : 1;

        var regularHours = hasDoneRegularHours ? worksheet.Cell(employeeRow + 1, dayColumn).GetString() : "";
        var overtimeHours = hasDoneOvertimeHours ? worksheet.Cell(employeeRow + overtimeIncrement, dayColumn).GetString() : "";
        return new TeamEmployeeWorkDayHours
        {
            RegularHours = TimeParsingUtils.ConvertToRoundedHours(regularHours),
            OvertimeHours = TimeParsingUtils.ConvertToRoundedHours(overtimeHours)
        };
    }

    public (decimal regularHoursRate, decimal overtimeRate) GetEmployeeRates(IXLWorksheet worksheet, int employeeRow, int ratesColumn)
    {
        var (hasDoneRegularHours, hasDoneOvertimeHours) = GetTypesOfHoursEmployeeHasDone(worksheet, employeeRow);
        var overtimeIncrement = hasDoneRegularHours ? 2 : 1;

        var regularHourRateStr = hasDoneRegularHours ? worksheet.Cell(employeeRow + 1, ratesColumn).GetString() : "";
        var overtimeRateStr = hasDoneOvertimeHours ? worksheet.Cell(employeeRow + overtimeIncrement, ratesColumn).GetString() : "";

        var regularHoursRate = decimal.TryParse(regularHourRateStr, out var parsedRegHrResult) ? parsedRegHrResult : 0;
        var overtimeRate = decimal.TryParse(overtimeRateStr, out var parsedOtHrResult) ? parsedOtHrResult : 0;

        return (regularHoursRate, overtimeRate);
    }

    public (bool hasRegularHours, bool hasOTHours) GetTypesOfHoursEmployeeHasDone(IXLWorksheet worksheet, int employeeRow)
    {
        var hourTypeCell = worksheet.FindSingleCellByValue(TeamSourceConstants.TypeCellValue);
        var hasRegularHours = worksheet.Cell(employeeRow + 1, hourTypeCell.Address.ColumnNumber).GetString().ToUpper() == TeamSourceConstants.RegularHours.ToUpper();
        var incrementForOt = hasRegularHours ? 2 : 1;
        var overtimeHours = worksheet.Cell(employeeRow + incrementForOt, hourTypeCell.Address.ColumnNumber).GetString().ToUpper() == TeamSourceConstants.WeeklyOt.ToUpper();
        return (hasRegularHours, overtimeHours);
    }
}

public interface ITeamSourceParser
{
    (DateOnly startDate, DateOnly endDate) GetStartAndEndDate(IXLWorksheet worksheet);
    TeamEmployeeWorkDayHours GetWorkdayHoursForEmployeeAndDay(IXLWorksheet worksheet, int employeeRow, int dayColumn);
    (decimal regularHoursRate, decimal overtimeRate) GetEmployeeRates(IXLWorksheet worksheet, int employeeRow, int ratesColumn);

    (bool hasRegularHours, bool hasOTHours) GetTypesOfHoursEmployeeHasDone(IXLWorksheet worksheet, int employeeRow);
}
