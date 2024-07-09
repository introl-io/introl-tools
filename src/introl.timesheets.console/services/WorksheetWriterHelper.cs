using ClosedXML.Excel;
using Introl.Timesheets.Console.constants;
using Introl.Timesheets.Console.models;

namespace Introl.Timesheets.Console.services;

public class WorksheetWriterHelper : IWorksheetWriterHelper
{
    private Dictionary<DayOfTheWeek, int> DayOfTheWeekColumnDictionary => new()
    {
        { DayOfTheWeek.Monday, 3 },
        { DayOfTheWeek.Tuesday, 4 },
        { DayOfTheWeek.Wednesday, 5 },
        { DayOfTheWeek.Thursday, 6 },
        { DayOfTheWeek.Friday, 7 },
        { DayOfTheWeek.Saturday, 8 },
        { DayOfTheWeek.Sunday, 9 },
    };

    private const int NameCol = 1;
    private const int HoursTypeCol = 2;
    private const int TotalHoursCol = 10;
    private const int RatesCol = 11;
    private const int TotalBillCol = 12;

    private const int BufferRow = 2;
    private const int WeekRow = 3;
    private const int DayRow = 4;
    private const int TitleRow = 5;
    
    public void AddTitleRows(IXLWorksheet worksheet, InputSheetModel inputSheetModel)
    {
        var weekRangedateFormat = "dd MMMM yyyy";
        var dayDateFormat = "MMMM dd";
        var pic = worksheet.AddPicture("./assets/introl_logo.png")
            .MoveTo(1, 1);
        pic.Width = DimensionConstants.ImageWightAndHeightInPixels;
        pic.Height = DimensionConstants.ImageWightAndHeightInPixels;
        
        for(var i = NameCol; i <= TotalBillCol; i++)
        {
            worksheet.Cell(BufferRow, i).Style.Fill.BackgroundColor = StyleConstants.Black;
        }
  
        var formattedDate = $"{inputSheetModel.StartDate.ToString(weekRangedateFormat)} - {inputSheetModel.EndDate.ToString(weekRangedateFormat)}";
        worksheet.Row(WeekRow).Style.Font.Bold = true;
        worksheet.Cell(WeekRow, 2).Value = formattedDate;
        
        worksheet.Row(DayRow).Style.Font.Bold = true;
        worksheet.Cell(DayRow, DayOfTheWeekColumnDictionary[DayOfTheWeek.Monday]).Value = "MON";
        worksheet.Cell(DayRow, DayOfTheWeekColumnDictionary[DayOfTheWeek.Tuesday]).Value = "TUE";
        worksheet.Cell(DayRow, DayOfTheWeekColumnDictionary[DayOfTheWeek.Wednesday]).Value = "WED";
        worksheet.Cell(DayRow, DayOfTheWeekColumnDictionary[DayOfTheWeek.Thursday]).Value = "THU";
        worksheet.Cell(DayRow, DayOfTheWeekColumnDictionary[DayOfTheWeek.Friday]).Value = "FRI";
        worksheet.Cell(DayRow, DayOfTheWeekColumnDictionary[DayOfTheWeek.Saturday]).Value = "SAT";
        worksheet.Cell(DayRow, DayOfTheWeekColumnDictionary[DayOfTheWeek.Sunday]).Value = "SUN";

        worksheet.Row(TitleRow).Style.Font.Bold = true;
        worksheet.Cell(TitleRow, NameCol).Value = "NAME";
        worksheet.Cell(TitleRow, HoursTypeCol).Value = "TYPE";
        for (var i = 0; i < 7; i++)
        {
            var mondayCol = DayOfTheWeekColumnDictionary[DayOfTheWeek.Monday];
            worksheet.Cell(TitleRow, mondayCol + i).Value = $"{inputSheetModel.StartDate.AddDays(i).ToString(dayDateFormat)}";
        }
        worksheet.Cell(TitleRow, TotalHoursCol).Value = "TOTALS";
        worksheet.Cell(TitleRow, RatesCol).Value = "RATES $";
        worksheet.Cell(TitleRow, TotalBillCol).Value = "TOTALS $";
    }
    
    public void AddEmployeeRows(IXLWorksheet worksheet, IEnumerable<Employee> employees, ref int employeeRow)
    {
        foreach (var employee in employees)
        {
            worksheet.Row(employeeRow).Style.Font.Bold = true;
            worksheet.Cell(employeeRow, NameCol).Value = employee.Name;
            worksheet.Cell(employeeRow, HoursTypeCol).Value = "Payroll Hours";
            worksheet.Cell(employeeRow+1, HoursTypeCol).Value = "Regular Hours";
            worksheet.Cell(employeeRow+2, HoursTypeCol).Value = "Weekly OT";

            foreach (var (dayOfTheWeek, col) in DayOfTheWeekColumnDictionary)
            {
                worksheet.Cell(employeeRow, col).Value = employee.WorkDays[dayOfTheWeek].TotalHours.ToString("F2");
                worksheet.Cell(employeeRow+1, col).Value = employee.WorkDays[dayOfTheWeek].RegularHours.ToString("F2");
                worksheet.Cell(employeeRow+2, col).Value = employee.WorkDays[dayOfTheWeek].OvertimeHours.ToString("F2");
            }
            
            worksheet.Cell(employeeRow, TotalHoursCol).Value = employee.TotalHours;
            worksheet.Cell(employeeRow+1, TotalHoursCol).Value = employee.TotalRegularHours;
            worksheet.Cell(employeeRow+2, TotalHoursCol).Value = employee.TotalOvertimeHours;
            
            worksheet.Cell(employeeRow+1, RatesCol).Value = employee.RegularHoursRate;
            worksheet.Cell(employeeRow+2, RatesCol).Value = employee.OvertimeHoursRate;
            
            worksheet.Cell(employeeRow, TotalBillCol).Value = employee.TotalBill;
            worksheet.Cell(employeeRow+1, TotalBillCol).Value = employee.TotalRegularBill;
            worksheet.Cell(employeeRow+2, TotalBillCol).Value = employee.TotalOvertimeBill;
            employeeRow += 3;
        }
    }

    public void AddTotals(IXLWorksheet worksheet, InputSheetModel inputSheetModel, int startRow)
    {
        var totalBillable = inputSheetModel.Employees.Sum(e => e.TotalBill);
        
        worksheet.Row(startRow).Style.Font.FontSize = StyleConstants.LargeFontSize;
        worksheet.Cell(startRow, HoursTypeCol).Value = "Total Hours";
        worksheet.Cell(startRow, RatesCol).Value = "Total $";
        worksheet.Cell(startRow, TotalBillCol).Value = totalBillable;
        
        worksheet.Cell(startRow + 1, HoursTypeCol).Value = "Payroll";
        worksheet.Cell(startRow + 2, HoursTypeCol).Value = "Regular";
        worksheet.Cell(startRow + 3, HoursTypeCol).Value = "Weekly OT";
        
        foreach (var (dayOfTheWeek, col) in DayOfTheWeekColumnDictionary)
        {
            var totalRegularHours = inputSheetModel.Employees.Sum(e => e.WorkDays[dayOfTheWeek].RegularHours);
            var totalOverTimeHours = inputSheetModel.Employees.Sum(e => e.WorkDays[dayOfTheWeek].OvertimeHours);
            var totalHours = totalRegularHours + totalOverTimeHours;
            
            worksheet.Cell(startRow + 1, col).Value = totalHours;
            worksheet.Cell(startRow + 2, col).Value = totalRegularHours;
            worksheet.Cell(startRow + 3, col).Value = totalOverTimeHours;
        }

        var weeksTotalRegularHours = inputSheetModel.Employees.Sum(e => e.TotalRegularHours);
        var weeksTotalOverTimeHours = inputSheetModel.Employees.Sum(e => e.TotalOvertimeHours);
        var weeksTotalHours = weeksTotalRegularHours + weeksTotalOverTimeHours;
        
        worksheet.Cell(startRow + 1, TotalHoursCol).Value = weeksTotalHours;
        worksheet.Cell(startRow + 2, TotalHoursCol).Value = weeksTotalRegularHours;
        worksheet.Cell(startRow + 3, TotalHoursCol).Value = weeksTotalOverTimeHours;

    }
}

public interface IWorksheetWriterHelper
{
    void AddTitleRows(IXLWorksheet worksheet, InputSheetModel inputSheetModel);
    void AddEmployeeRows(IXLWorksheet worksheet, IEnumerable<Employee> employees, ref int employeeRow);
    void AddTotals(IXLWorksheet worksheet, InputSheetModel inputSheetModel, int startRow);
}
