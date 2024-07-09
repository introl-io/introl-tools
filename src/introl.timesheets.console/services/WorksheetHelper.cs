using ClosedXML.Excel;
using Introl.Timesheets.Console.models;

namespace Introl.Timesheets.Console.services;

public class WorksheetHelper : IWorksheetHelper
{
    public IXLCell FindSingleCellByValue(IXLWorksheet worksheet, string value)
    {
        var matchingCells = worksheet.CellsUsed(c => c.GetString().ToUpper() == value.ToUpper());
        if(!matchingCells.Any())
        {
            throw new ArgumentNullException($"No cell found with the value {value}");
        }
        
        if(matchingCells.Count() > 1)
        {
            throw new InvalidOperationException($"Multiple cells found with the value {value}");
        }
        return matchingCells.First();
    }

    public IDictionary<DayOfTheWeek, int> GetDayOfTheWeekColumnDictionary(IXLWorksheet worksheet)
    {
        var monCol = FindSingleCellByValue(worksheet, "mon").Address.ColumnNumber;
        var tueCol = FindSingleCellByValue(worksheet, "tue").Address.ColumnNumber;
        var wedCol = FindSingleCellByValue(worksheet, "wed").Address.ColumnNumber;
        var thuCol = FindSingleCellByValue(worksheet, "thu").Address.ColumnNumber;
        var friCol = FindSingleCellByValue(worksheet, "fri").Address.ColumnNumber;
        var satCol = FindSingleCellByValue(worksheet, "sat").Address.ColumnNumber;
        var sunCol = FindSingleCellByValue(worksheet, "sun").Address.ColumnNumber;

        return new Dictionary<DayOfTheWeek, int>
        {
            { DayOfTheWeek.Monday, monCol },
            { DayOfTheWeek.Tuesday, tueCol },
            { DayOfTheWeek.Wednesday, wedCol },
            { DayOfTheWeek.Thursday, thuCol },
            { DayOfTheWeek.Friday, friCol },
            { DayOfTheWeek.Saturday, satCol },
            { DayOfTheWeek.Sunday, sunCol },
        };
    }

    public (DateOnly startDate, DateOnly endDate) GetStartAndEndDate(IXLWorksheet worksheet)
    {
        var weekCell = FindSingleCellByValue(worksheet, "week");
        var dateString = weekCell.CellRight().GetString();
        var splitDates = dateString.Split(" - ");
        
        var startDate = DateOnly.Parse(splitDates[0]);
        var endDate = DateOnly.Parse(splitDates[1]);

        return (startDate, endDate);
    }
    
    public WorkDayHours GetWorkdayHoursForEmployeeAndDay(IXLWorksheet worksheet, int employeeRow, int dayColumn)
    {
        var regularHours = worksheet.Cell(employeeRow + 1, dayColumn).GetString();
        var overtimeHours = worksheet.Cell(employeeRow + 2, dayColumn).GetString();
        return new WorkDayHours
        {
            RegularHours = ConvertToRoundedHours(regularHours),
            OvertimeHours = ConvertToRoundedHours(overtimeHours)
        };
    }

    public (decimal regularHoursRate, decimal overtimeRate) GetEmployeeRates(IXLWorksheet worksheet, int employeeRow, int ratesColumn)
    {
        var regularHourRateStr = worksheet.Cell(employeeRow + 1, ratesColumn).GetString();
        var overtimeRateStr = worksheet.Cell(employeeRow + 2, ratesColumn).GetString();
        var regularHoursRate = !string.IsNullOrEmpty(regularHourRateStr) ? decimal.Parse(regularHourRateStr) : 0;
        var overtimeRate = !string.IsNullOrEmpty(overtimeRateStr) ? decimal.Parse(overtimeRateStr) : 0;

        return (regularHoursRate, overtimeRate);
    }

    private double ConvertToRoundedHours(string inputHours)
    {
        if (!inputHours.Contains(":"))
        {
            if(double.TryParse(inputHours, out var parsedHours))
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

public interface IWorksheetHelper
{
    IXLCell FindSingleCellByValue(IXLWorksheet worksheet, string value);
    IDictionary<DayOfTheWeek, int> GetDayOfTheWeekColumnDictionary(IXLWorksheet worksheet);
    (DateOnly startDate, DateOnly endDate) GetStartAndEndDate(IXLWorksheet worksheet);
    WorkDayHours GetWorkdayHoursForEmployeeAndDay(IXLWorksheet worksheet, int employeeRow, int dayColumn);
    (decimal regularHoursRate, decimal overtimeRate) GetEmployeeRates(IXLWorksheet worksheet, int employeeRow, int ratesColumn);
}
