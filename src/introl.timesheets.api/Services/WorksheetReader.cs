using ClosedXML.Excel;
using Introl.Timesheets.Api.models;

namespace Introl.Timesheets.Api.Services;

public class WorksheetReader(IWorksheetReaderHelper worksheetReaderHelper) : IWorksheetReader
{
    public InputSheetModel Process(XLWorkbook workbook)
    {
        var teamSummarySheet = workbook.Worksheets.Worksheet("Team Summary");
        var rawTimesheetsWorkSheet = workbook.Worksheets.Worksheet("Raw Timesheets");
        ArgumentNullException.ThrowIfNull(teamSummarySheet);
        var (startDate, endDate) = worksheetReaderHelper.GetStartAndEndDate(teamSummarySheet);
        
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
        var cell = worksheetReaderHelper.FindSingleCellByValue(worksheet, "name");
        return cell.Address.RowNumber + 1;
    }
    
    private int RatesColumn(IXLWorksheet worksheet)
    {
        var cell = worksheetReaderHelper.FindSingleCellByValue(worksheet, "RATES");
        return cell.Address.ColumnNumber;
    }

    private IEnumerable<Employee> GetEmployees(IXLWorksheet worksheet)
    {
        var employeeRow = GetFirstEmployeeRow(worksheet);
        var dayColDict = worksheetReaderHelper.GetDayOfTheWeekColumnDictionary(worksheet);
        var ratesCol = RatesColumn(worksheet);
        do
        {
            yield return GetEmployee(worksheet, employeeRow, dayColDict, ratesCol);
            employeeRow += 3;
        } while (!string.IsNullOrEmpty(worksheet.Cell(employeeRow, 1).GetString()));
    }
    
    private Employee GetEmployee(IXLWorksheet worksheet, int employeeRow, IDictionary<DayOfTheWeek, int> dayDictionary, int ratesCol)
    {
        var name = worksheet.Cell(employeeRow, 1).GetString();
        var (regularHoursRate, overtimeHoursRate) = worksheetReaderHelper.GetEmployeeRates(worksheet, employeeRow, ratesCol);
        
        return new Employee
        {
            Name = name,
            RegularHoursRate = regularHoursRate,
            OvertimeHoursRate = overtimeHoursRate,
            WorkDays = new Dictionary<DayOfTheWeek, WorkDayHours>
            {
                { DayOfTheWeek.Monday, worksheetReaderHelper.GetWorkdayHoursForEmployeeAndDay(worksheet, employeeRow, dayDictionary[DayOfTheWeek.Monday]) },
                { DayOfTheWeek.Tuesday, worksheetReaderHelper.GetWorkdayHoursForEmployeeAndDay(worksheet, employeeRow, dayDictionary[DayOfTheWeek.Tuesday]) },
                { DayOfTheWeek.Wednesday, worksheetReaderHelper.GetWorkdayHoursForEmployeeAndDay(worksheet, employeeRow, dayDictionary[DayOfTheWeek.Wednesday]) },
                { DayOfTheWeek.Thursday, worksheetReaderHelper.GetWorkdayHoursForEmployeeAndDay(worksheet, employeeRow, dayDictionary[DayOfTheWeek.Thursday]) },
                { DayOfTheWeek.Friday, worksheetReaderHelper.GetWorkdayHoursForEmployeeAndDay(worksheet, employeeRow, dayDictionary[DayOfTheWeek.Friday]) },
                { DayOfTheWeek.Saturday, worksheetReaderHelper.GetWorkdayHoursForEmployeeAndDay(worksheet, employeeRow, dayDictionary[DayOfTheWeek.Saturday]) },
                { DayOfTheWeek.Sunday, worksheetReaderHelper.GetWorkdayHoursForEmployeeAndDay(worksheet, employeeRow, dayDictionary[DayOfTheWeek.Sunday]) }
            }
        };
    }
}

public interface IWorksheetReader
{
    InputSheetModel Process(XLWorkbook workbook);
}
