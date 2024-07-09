using ClosedXML.Excel;
using Introl.Timesheets.Console.models;

namespace Introl.Timesheets.Console.services;

public class WorksheetWriter : IWorksheetWriter
{
    public void Process(InputSheetModel inputSheetModel)
    {
        using var workbook = new XLWorkbook();

        CreateSummarySheet(workbook, inputSheetModel);
        
        workbook.AddWorksheet(inputSheetModel.RawTimesheetsWorksheet);
        workbook.SaveAs("./output/output.xlsx");
    }

    private void CreateSummarySheet(XLWorkbook workbook, InputSheetModel inputSheetModel)
    {
        var worksheet = workbook.Worksheets.Add("Summary");
        AddTitleRows(worksheet, inputSheetModel);
        AddEmployeeRows(worksheet, inputSheetModel.Employees);
    }

    private void AddTitleRows(IXLWorksheet worksheet, InputSheetModel inputSheetModel)
    {
        var weekRangedateFormat = "dd MMMM yyyy";
        var dayDateFormat = "MMMM dd";

        worksheet.Cell(2, 2).Value =
            $"{inputSheetModel.StartDate.ToString(weekRangedateFormat)} - {inputSheetModel.EndDate.ToString(weekRangedateFormat)}";

        worksheet.Cell(3, 3).Value = "MON";
        worksheet.Cell(3, 4).Value = "TUE";
        worksheet.Cell(3, 5).Value = "WED";
        worksheet.Cell(3, 6).Value = "THU";
        worksheet.Cell(3, 7).Value = "FRI";
        worksheet.Cell(3, 8).Value = "SAT";
        worksheet.Cell(3, 9).Value = "SUN";

        worksheet.Cell(4, 1).Value = "NAME";
        worksheet.Cell(4, 2).Value = "TYPE";
        for (var i = 0; i < 7; i++)
        {
            worksheet.Cell(4, 3 + i).Value = $"{inputSheetModel.StartDate.AddDays(i).ToString(dayDateFormat)}";
        }
        worksheet.Cell(4, 10).Value = "TOTALS";
        worksheet.Cell(4, 11).Value = "RATES $";
        worksheet.Cell(4, 12).Value = "TOTALS $";
    }
    
    private void AddEmployeeRows(IXLWorksheet worksheet, IEnumerable<Employee> employees)
    {
        var row = 5;
        foreach (var employee in employees)
        {
            var totalRegularHours = employee.WorkDays.Sum(w => w.Value.RegularHours);
            var totalOvertimeHours = employee.WorkDays.Sum(w => w.Value.OvertimeHours);
            
            var totalRegularBill = (decimal)totalRegularHours * employee.RegularHoursRate;
            var totalOvertimeBill = (decimal)totalOvertimeHours * employee.OvertimeHoursRate;
            
            worksheet.Cell(row, 1).Value = employee.Name;
            worksheet.Cell(row, 2).Value = "Payroll Hours";
            worksheet.Cell(row+1, 2).Value = "Regular Hours";
            worksheet.Cell(row+2, 2).Value = "Weekly OT";
            
            worksheet.Cell(row, 3).Value = employee.WorkDays[DayOfTheWeek.Monday].TotalHours;
            worksheet.Cell(row+1, 3).Value = employee.WorkDays[DayOfTheWeek.Monday].RegularHours;
            worksheet.Cell(row+2, 3).Value = employee.WorkDays[DayOfTheWeek.Monday].OvertimeHours;
            
            worksheet.Cell(row, 4).Value = employee.WorkDays[DayOfTheWeek.Tuesday].TotalHours;
            worksheet.Cell(row+1, 4).Value = employee.WorkDays[DayOfTheWeek.Tuesday].RegularHours;
            worksheet.Cell(row+2, 4).Value = employee.WorkDays[DayOfTheWeek.Tuesday].OvertimeHours;
            
            worksheet.Cell(row, 5).Value = employee.WorkDays[DayOfTheWeek.Wednesday].TotalHours;
            worksheet.Cell(row+1, 5).Value = employee.WorkDays[DayOfTheWeek.Wednesday].RegularHours;
            worksheet.Cell(row+2, 5).Value = employee.WorkDays[DayOfTheWeek.Wednesday].OvertimeHours;
            
            worksheet.Cell(row, 6).Value = employee.WorkDays[DayOfTheWeek.Thursday].TotalHours;
            worksheet.Cell(row+1, 6).Value = employee.WorkDays[DayOfTheWeek.Thursday].RegularHours;
            worksheet.Cell(row+2, 6).Value = employee.WorkDays[DayOfTheWeek.Thursday].OvertimeHours;
            
            worksheet.Cell(row, 7).Value = employee.WorkDays[DayOfTheWeek.Friday].TotalHours;
            worksheet.Cell(row+1, 7).Value = employee.WorkDays[DayOfTheWeek.Friday].RegularHours;
            worksheet.Cell(row+2, 7).Value = employee.WorkDays[DayOfTheWeek.Friday].OvertimeHours;
            
            worksheet.Cell(row, 8).Value = employee.WorkDays[DayOfTheWeek.Saturday].TotalHours;
            worksheet.Cell(row+1, 8).Value = employee.WorkDays[DayOfTheWeek.Saturday].RegularHours;
            worksheet.Cell(row+2, 8).Value = employee.WorkDays[DayOfTheWeek.Saturday].OvertimeHours;
            
            worksheet.Cell(row, 9).Value = employee.WorkDays[DayOfTheWeek.Sunday].TotalHours;
            worksheet.Cell(row+1, 9).Value = employee.WorkDays[DayOfTheWeek.Sunday].RegularHours;
            worksheet.Cell(row+2, 9).Value = employee.WorkDays[DayOfTheWeek.Sunday].OvertimeHours;

            worksheet.Cell(row, 10).Value = totalRegularHours + totalOvertimeHours;
            worksheet.Cell(row+1, 10).Value = totalRegularHours;
            worksheet.Cell(row+2, 10).Value = totalOvertimeHours;
            
            worksheet.Cell(row+1, 11).Value = employee.RegularHoursRate;
            worksheet.Cell(row+2, 11).Value = employee.OvertimeHoursRate;
            
            worksheet.Cell(row, 12).Value = totalRegularBill + totalOvertimeBill;
            worksheet.Cell(row+1, 12).Value = totalRegularBill;
            worksheet.Cell(row+2, 12).Value = totalOvertimeBill;
            row += 3;
        }
    }
}

public interface IWorksheetWriter
{
    void Process(InputSheetModel inputSheetModel);
}
