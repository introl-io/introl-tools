using ClosedXML.Excel;
using Introl.Timesheets.Console.models;

namespace Introl.Timesheets.Console.services;

public class WorksheetReader : IWorksheetReader
{
    public InputSheetModel Process(XLWorkbook workbook)
    {
        var teamSummarySheet = workbook.Worksheets.Worksheet("Team Summary");
        var rawTimesheetsWorkSheet = workbook.Worksheets.Worksheet("Raw Timesheets");
        ArgumentNullException.ThrowIfNull(teamSummarySheet);
        var (startDate, endDate) = GetDates(teamSummarySheet);
        
        return new InputSheetModel
        {
            StartDate = startDate,
            EndDate = endDate,
            Employees = GetEmployees(teamSummarySheet),
            RawTimesheetsWorksheet = rawTimesheetsWorkSheet
        };
    }

    private int GetFirstEmployeeRow(IXLWorksheet worksheet)
    {
        var colA = worksheet.Column("A")!;
        var cells = colA.CellsUsed(c => c.GetString().ToUpper() == "NAME");
        return cells.ToList().First().Address.RowNumber + 1;
    }

    private int GetMondayColumn(IXLWorksheet worksheet)
    {
        var mondayCell = worksheet.CellsUsed(c => c.GetString().ToUpper() == "MON");
        ArgumentNullException.ThrowIfNull(mondayCell);

        return mondayCell.First().Address.ColumnNumber;
    }

    private int RatesColumn(IXLWorksheet worksheet)
    {
        var mondayCell = worksheet.CellsUsed(c => c.GetString().ToUpper() == "RATES");
        ArgumentNullException.ThrowIfNull(mondayCell);

        return mondayCell.First().Address.ColumnNumber;
    }

    private IEnumerable<Employee> GetEmployees(IXLWorksheet worksheet)
    {
        var employeeRow = GetFirstEmployeeRow(worksheet);
        var monCol = GetMondayColumn(worksheet);
        var ratesCol = RatesColumn(worksheet);
        do
        {
            yield return GetEmployee(worksheet, employeeRow, monCol, ratesCol);
            employeeRow += 3;
        } while (!string.IsNullOrEmpty(worksheet.Cell(employeeRow, 1).GetString()));
    }
    
    private Employee GetEmployee(IXLWorksheet worksheet, int employeeRow, int mondayCol, int ratesCol)
    {
        var name = worksheet.Cell(employeeRow, 1).GetString();
        var regularHourRateStr = worksheet.Cell(employeeRow + 1, ratesCol).GetString();
        var overtimeHoursRateStr = worksheet.Cell(employeeRow + 2, ratesCol).GetString();
        var regularHoursRate = !string.IsNullOrEmpty(regularHourRateStr) ? decimal.Parse(regularHourRateStr) : 0;
        var overtimeHoursRate = !string.IsNullOrEmpty(overtimeHoursRateStr) ? decimal.Parse(overtimeHoursRateStr) : 0;
        return new Employee
        {
            Name = name,
            RegularHoursRate = regularHoursRate,
            OvertimeHoursRate = overtimeHoursRate,
            WorkDays = new Dictionary<DayOfTheWeek, WorkDayHours>
            {
                { DayOfTheWeek.Monday, GetWorkDayHoursForDay(worksheet, employeeRow, mondayCol) },
                { DayOfTheWeek.Tuesday, GetWorkDayHoursForDay(worksheet, employeeRow, mondayCol + 1) },
                { DayOfTheWeek.Wednesday, GetWorkDayHoursForDay(worksheet, employeeRow, mondayCol + 2) },
                { DayOfTheWeek.Thursday, GetWorkDayHoursForDay(worksheet, employeeRow, mondayCol + 3) },
                { DayOfTheWeek.Friday, GetWorkDayHoursForDay(worksheet, employeeRow, mondayCol + 4) },
                { DayOfTheWeek.Saturday, GetWorkDayHoursForDay(worksheet, employeeRow, mondayCol + 5) },
                { DayOfTheWeek.Sunday, GetWorkDayHoursForDay(worksheet, employeeRow, mondayCol + 6) }
            }
        };
    }

    private (DateOnly startDate, DateOnly endDate) GetDates(IXLWorksheet worksheet)
    {
        var weekRow = worksheet.CellsUsed(c => c.GetString().ToUpper() == "WEEK");
        var dateString = weekRow.First().CellRight().GetString();
        var splitDates = dateString.Split(" - ");
        
        var startDate = DateOnly.Parse(splitDates[0]);
        var endDate = DateOnly.Parse(splitDates[1]);

        return (startDate, endDate);
    }

    private WorkDayHours GetWorkDayHoursForDay(IXLWorksheet worksheet, int employeeRow, int dayCol)
    {
        var regularHoursInput = worksheet.Cell(employeeRow + 1, dayCol).GetString();
        var overtimeHoursInput = worksheet.Cell(employeeRow + 2, dayCol).GetString();
        return new WorkDayHours
        {
            RegularHours = ConvertToRoundedHours(regularHoursInput),
            OvertimeHours = ConvertToRoundedHours(overtimeHoursInput)
        };
    }

    private double ConvertToRoundedHours(string inputHours)
    {
        if (!inputHours.Contains(":"))
        {
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

public interface IWorksheetReader
{
    InputSheetModel Process(XLWorkbook workbook);
}
