using ClosedXML.Excel;
using Introl.Timesheets.Api.Constants;
using Introl.Timesheets.Api.Enums;
using Introl.Timesheets.Api.Extensions;
using Introl.Timesheets.Api.Models;

namespace Introl.Timesheets.Api.Services;

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

    private const int NameColInt = 1;
    private const int HoursTypeColInt = 2;
    private const int TotalHoursColInt = 10;
    private const int RatesColInt = 11;
    private const int TotalBillColInt = 12;

    private string HoursTypeColLetter => HoursTypeColInt.ToExcelColumn();
    private string TotalHoursColLetter => TotalHoursColInt.ToExcelColumn();
    private string RatesColLetter => RatesColInt.ToExcelColumn();
    private string TotalBillColLetter => TotalBillColInt.ToExcelColumn();
    
    private const int BufferRow = 2;
    private const int WeekRow = 3;
    private const int DayRow = 4;
    private const int TitleRow = 5;

    public void AddTitleRows(IXLWorksheet worksheet, InputSheetModel inputSheetModel)
    {
        var pic = worksheet.AddPicture("./Assets/introl_logo.png")
            .MoveTo(1, 1);
        pic.Width = DimensionConstants.ImageWightAndHeightInPixels;
        pic.Height = DimensionConstants.ImageWightAndHeightInPixels;

        for (var i = NameColInt; i <= TotalBillColInt; i++)
        {
            worksheet.Cell(BufferRow, i).Style.Fill.BackgroundColor = StyleConstants.Black;
        }

        AddWeekRow(worksheet, inputSheetModel);
        AddDayRow(worksheet);
        AddTitleRow(worksheet, inputSheetModel);
    }

    private void AddTitleRow(IXLWorksheet worksheet, InputSheetModel inputSheetModel)
    {
        var dayDateFormat = "MMMM dd";

        worksheet.Row(TitleRow).Style.Font.Bold = true;
        worksheet.Cell(TitleRow, NameColInt).Value = "NAME";
        worksheet.Cell(TitleRow, NameColInt).Style.Fill.BackgroundColor = StyleConstants.DarkGrey;

        worksheet.Cell(TitleRow, HoursTypeColInt).Value = "TYPE";
        worksheet.Cell(TitleRow, HoursTypeColInt).Style.Fill.BackgroundColor = StyleConstants.DarkGrey;
        for (var i = 0; i < 7; i++)
        {
            var mondayCol = DayOfTheWeekColumnDictionary[DayOfTheWeek.Monday];
            worksheet.Cell(TitleRow, mondayCol + i).Value =
                $"{inputSheetModel.StartDate.AddDays(i).ToString(dayDateFormat)}";
            worksheet.Cell(TitleRow, mondayCol + i).Style.Fill.BackgroundColor = StyleConstants.DarkGrey;
        }

        worksheet.Cell(TitleRow, TotalHoursColInt).Value = "TOTALS";
        worksheet.Cell(TitleRow, TotalHoursColInt).Style.Fill.BackgroundColor = StyleConstants.DarkGrey;

        worksheet.Cell(TitleRow, RatesColInt).Value = "RATES $";
        worksheet.Cell(TitleRow, RatesColInt).Style.Fill.BackgroundColor = StyleConstants.DarkGrey;
        worksheet.Cell(TitleRow, TotalBillColInt).Value = "TOTALS $";

        worksheet.Cell(TitleRow, TotalBillColInt).Style.Fill.BackgroundColor = StyleConstants.DarkGrey;
    }

    private void AddDayRow(IXLWorksheet worksheet)
    {
        worksheet.Row(DayRow).Style.Font.Bold = true;
        foreach (var day in Enum.GetValues(typeof(DayOfTheWeek)).Cast<DayOfTheWeek>())
        {
            worksheet.Cell(DayRow, DayOfTheWeekColumnDictionary[day]).Style.Fill.BackgroundColor =
                StyleConstants.LightGrey;
            worksheet.Cell(DayRow, DayOfTheWeekColumnDictionary[day]).Value = day.StringValue().ToUpper();
        }
    }

    private void AddWeekRow(IXLWorksheet worksheet, InputSheetModel inputSheetModel)
    {
        var weekRangedateFormat = "dd MMMM yyyy";
        var formattedDate =
            $"{inputSheetModel.StartDate.ToString(weekRangedateFormat)} - {inputSheetModel.EndDate.ToString(weekRangedateFormat)}";
        worksheet.Row(WeekRow).Style.Font.Bold = true;
        worksheet.Cell(WeekRow, 2).Value = formattedDate;
    }

    public void AddEmployeeRows(IXLWorksheet worksheet, IEnumerable<Employee> employees, ref int employeeRow)
    {
        var mondayColLetter = DayOfTheWeekColumnDictionary[DayOfTheWeek.Monday].ToExcelColumn();
        var sundayColLetter = DayOfTheWeekColumnDictionary[DayOfTheWeek.Sunday].ToExcelColumn();
      
        foreach (var employee in employees)
        {
            worksheet.Row(employeeRow).Style.Font.Bold = true;
            worksheet.Cell(employeeRow, NameColInt).Value = employee.Name;
            worksheet.Cell(employeeRow, HoursTypeColInt).Value = "Payroll Hours";
            worksheet.Cell(employeeRow + 1, HoursTypeColInt).Value = "Regular Hours";
            worksheet.Cell(employeeRow + 2, HoursTypeColInt).Value = "Weekly OT";

            foreach (var (dayOfTheWeek, col) in DayOfTheWeekColumnDictionary)
            {
                worksheet.Cell(employeeRow, col).Value = employee.WorkDays[dayOfTheWeek].TotalHours;
                worksheet.Cell(employeeRow, col).Style.NumberFormat.Format = StyleConstants.HourCellFormat;

                worksheet.Cell(employeeRow + 1, col).Value = employee.WorkDays[dayOfTheWeek].RegularHours;
                worksheet.Cell(employeeRow + 1, col).Style.NumberFormat.Format = StyleConstants.HourCellFormat;

                worksheet.Cell(employeeRow + 2, col).Value = employee.WorkDays[dayOfTheWeek].OvertimeHours;
                worksheet.Cell(employeeRow + 2, col).Style.NumberFormat.Format = StyleConstants.HourCellFormat;
            }

            worksheet.Cell(employeeRow, TotalHoursColInt).FormulaA1 =
                $"SUM({mondayColLetter}{employeeRow}:{sundayColLetter}{employeeRow})";
            worksheet.Cell(employeeRow, TotalHoursColInt).Style.NumberFormat.Format = StyleConstants.HourCellFormat;

            worksheet.Cell(employeeRow + 1, TotalHoursColInt).FormulaA1 =
                $"SUM({mondayColLetter}{employeeRow + 1}:{sundayColLetter}{employeeRow + 1})";
            worksheet.Cell(employeeRow + 1, TotalHoursColInt).Style.NumberFormat.Format = StyleConstants.HourCellFormat;

            worksheet.Cell(employeeRow + 2, TotalHoursColInt).FormulaA1 =
                $"SUM({mondayColLetter}{employeeRow + 2}:{sundayColLetter}{employeeRow + 2})";
            worksheet.Cell(employeeRow + 2, TotalHoursColInt).Style.NumberFormat.Format = StyleConstants.HourCellFormat;

            worksheet.Cell(employeeRow + 1, RatesColInt).Value = employee.RegularHoursRate;
            worksheet.Cell(employeeRow + 1, RatesColInt).Style.NumberFormat.Format = StyleConstants.CurrencyCellFormat;
            worksheet.Cell(employeeRow + 2, RatesColInt).Value = employee.OvertimeHoursRate;
            worksheet.Cell(employeeRow + 2, RatesColInt).Style.NumberFormat.Format = StyleConstants.CurrencyCellFormat;

            worksheet.Cell(employeeRow, TotalBillColInt).FormulaA1 =
                $"{TotalBillColLetter}{employeeRow + 1} + {TotalBillColLetter}{employeeRow + 2}";
            worksheet.Cell(employeeRow, TotalBillColInt).Style.NumberFormat.Format = StyleConstants.CurrencyCellFormat;

            worksheet.Cell(employeeRow + 1, TotalBillColInt).FormulaA1 =
                $"{TotalHoursColLetter}{employeeRow + 1} * {RatesColLetter}{employeeRow + 1}";
            ;
            worksheet.Cell(employeeRow + 1, TotalBillColInt).Style.NumberFormat.Format = StyleConstants.CurrencyCellFormat;

            worksheet.Cell(employeeRow + 2, TotalBillColInt).FormulaA1 =
                $"{TotalHoursColLetter}{employeeRow + 2} * {RatesColLetter}{employeeRow + 2}";
            ;
            worksheet.Cell(employeeRow + 2, TotalBillColInt).Style.NumberFormat.Format = StyleConstants.CurrencyCellFormat;
            employeeRow += 3;
        }
    }

    public void AddTotals(IXLWorksheet worksheet, InputSheetModel inputSheetModel, int totalsStartRow,
        int lastEmployeeRow)
    {
        var typeRange = $"{HoursTypeColLetter}1:{HoursTypeColLetter}{lastEmployeeRow}";
        var totalBillableRange = $"{TotalBillColLetter}1:{TotalBillColLetter}{lastEmployeeRow}";
        worksheet.Row(totalsStartRow).Style.Font.Bold = true;
        worksheet.Row(totalsStartRow).Style.Font.FontSize = StyleConstants.LargeFontSize;
        worksheet.Cell(totalsStartRow, HoursTypeColInt).Value = "Total Hours";
        worksheet.Cell(totalsStartRow, RatesColInt).Value = "Total $";
        worksheet.Cell(totalsStartRow, TotalBillColInt).FormulaA1 = $"=SUMIFS({totalBillableRange},{typeRange},\"Payroll Hours\")";
        ;
        worksheet.Cell(totalsStartRow, TotalBillColInt).Style.NumberFormat.Format =
            StyleConstants.CurrencyWithSymbolCellFormat;

        worksheet.Cell(totalsStartRow + 1, HoursTypeColInt).Value = "Payroll";
        worksheet.Cell(totalsStartRow + 2, HoursTypeColInt).Value = "Regular";
        worksheet.Cell(totalsStartRow + 3, HoursTypeColInt).Value = "Weekly OT";

        foreach (var (_, col) in DayOfTheWeekColumnDictionary)
        {
            var colLetter = col.ToExcelColumn();
            var daysHourRangeRange = $"{colLetter}1:{colLetter}{lastEmployeeRow}";
            
            worksheet.Cell(totalsStartRow + 1, col).FormulaA1 =
                $"{colLetter}{totalsStartRow + 2} + {colLetter}{totalsStartRow + 3}";
            worksheet.Cell(totalsStartRow, col).Style.NumberFormat.Format = StyleConstants.HourCellFormat;

            worksheet.Cell(totalsStartRow + 2, col).FormulaA1 = $"=SUMIFS({daysHourRangeRange},{typeRange},\"Regular Hours\")";
            worksheet.Cell(totalsStartRow + 2, col).Style.NumberFormat.Format = StyleConstants.HourCellFormat;

            worksheet.Cell(totalsStartRow + 3, col).FormulaA1 = $"=SUMIFS({daysHourRangeRange},{typeRange},\"Weekly OT\")";
            worksheet.Cell(totalsStartRow + 3, col).Style.NumberFormat.Format = StyleConstants.HourCellFormat;
        }

        var mondayColLetter = DayOfTheWeekColumnDictionary[DayOfTheWeek.Monday].ToExcelColumn();
        var sundayColLetter = DayOfTheWeekColumnDictionary[DayOfTheWeek.Sunday].ToExcelColumn();
        worksheet.Cell(totalsStartRow + 1, TotalHoursColInt).FormulaA1 = $"SUM({mondayColLetter}{totalsStartRow + 1}:{sundayColLetter}{totalsStartRow + 1})";
        worksheet.Cell(totalsStartRow + 2, TotalHoursColInt).FormulaA1 = $"SUM({mondayColLetter}{totalsStartRow + 2}:{sundayColLetter}{totalsStartRow + 2})";
        worksheet.Cell(totalsStartRow + 3, TotalHoursColInt).FormulaA1 = $"SUM({mondayColLetter}{totalsStartRow + 3}:{sundayColLetter}{totalsStartRow + 3})";
    }
}

public interface IWorksheetWriterHelper
{
    void AddTitleRows(IXLWorksheet worksheet, InputSheetModel inputSheetModel);
    void AddEmployeeRows(IXLWorksheet worksheet, IEnumerable<Employee> employees, ref int employeeRow);
    void AddTotals(IXLWorksheet worksheet, InputSheetModel inputSheetModel, int totalsStartRow, int lastEmployeeRow);
}
