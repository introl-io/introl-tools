using ClosedXML.Excel;
using Introl.Timesheets.Api.Enums;
using Introl.Timesheets.Api.Extensions;
using Introl.Timesheets.Api.Models;

namespace Introl.Timesheets.Api.Services;

public class WorksheetReaderHelper : IWorksheetReaderHelper
{
    public IDictionary<DayOfTheWeek, int> GetDayOfTheWeekColumnDictionary(IXLWorksheet worksheet)
    {
        var result = new Dictionary<DayOfTheWeek, int>();
        foreach (var day in Enum.GetValues(typeof(DayOfTheWeek)).Cast<DayOfTheWeek>())
        {
            result.Add(day, worksheet.FindSingleCellByValue(day.StringValue()).Address.ColumnNumber);
        }

        return result;
    }

    public (DateOnly startDate, DateOnly endDate) GetStartAndEndDate(IXLWorksheet worksheet)
    {
        var weekCell = worksheet.FindSingleCellByValue("week");
        var dateString = weekCell.CellRight().GetString();
        var splitDates = dateString.Split(" - ");

        var startDate = DateOnly.Parse(splitDates[0]);
        var endDate = DateOnly.Parse(splitDates[1]);

        return (startDate, endDate);
    }

    public WorkDayHours GetWorkdayHoursForEmployeeAndDay(IXLWorksheet worksheet, int employeeRow, int dayColumn)
    {

        var (hasDoneRegularHours, hasDoneOvertimeHours) = GetTypesOfHoursEmployeeHasDone(worksheet, employeeRow);
        var overtimeIncrement = hasDoneRegularHours ? 2 : 1;

        var regularHours = hasDoneRegularHours ? worksheet.Cell(employeeRow + 1, dayColumn).GetString() : "";
        var overtimeHours = hasDoneOvertimeHours ? worksheet.Cell(employeeRow + overtimeIncrement, dayColumn).GetString() : "";
        return new WorkDayHours
        {
            RegularHours = ConvertToRoundedHours(regularHours),
            OvertimeHours = ConvertToRoundedHours(overtimeHours)
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
        var hourTypeCell = worksheet.FindSingleCellByValue("type");
        var hasRegularHours = worksheet.Cell(employeeRow + 1, hourTypeCell.Address.ColumnNumber).GetString().ToUpper() == "REGULAR HOURS";
        var incrementForOt = hasRegularHours ? 2 : 1;
        var overtimeHours = worksheet.Cell(employeeRow + incrementForOt, hourTypeCell.Address.ColumnNumber).GetString().ToUpper() == "WEEKLY OT";
        return (hasRegularHours, overtimeHours);
    }

    private double ConvertToRoundedHours(string inputHours)
    {
        if (!inputHours.Contains(":"))
        {
            if (double.TryParse(inputHours, out var parsedHours))
            {
                return parsedHours;
            }

            return 0;
        }

        var splitHours = inputHours.Split(':');
        var hours = double.Parse(splitHours[0]);
        var minutes = double.Parse(splitHours[1]);

        return minutes switch
        {
            >= 0 and <= 14 => hours,
            >= 15 and <= 44 => hours + 0.5,
            _ => hours + 1
        };
    }
}

public interface IWorksheetReaderHelper
{
    IDictionary<DayOfTheWeek, int> GetDayOfTheWeekColumnDictionary(IXLWorksheet worksheet);
    (DateOnly startDate, DateOnly endDate) GetStartAndEndDate(IXLWorksheet worksheet);
    WorkDayHours GetWorkdayHoursForEmployeeAndDay(IXLWorksheet worksheet, int employeeRow, int dayColumn);
    (decimal regularHoursRate, decimal overtimeRate) GetEmployeeRates(IXLWorksheet worksheet, int employeeRow, int ratesColumn);

    (bool hasRegularHours, bool hasOTHours) GetTypesOfHoursEmployeeHasDone(IXLWorksheet worksheet, int employeeRow);
}
