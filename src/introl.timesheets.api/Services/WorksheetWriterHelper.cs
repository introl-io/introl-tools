using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
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

        worksheet.Row(TitleRow).Style.Font.Bold = true;
        worksheet.Row(DayRow).Style.Font.Bold = true;
        worksheet.Row(WeekRow).Style.Font.Bold = true;
        
        AddWeekRow(worksheet, inputSheetModel);
        AddDayRow(worksheet);
        AddTitleRow(worksheet, inputSheetModel);
    }

    private void AddTitleRow(IXLWorksheet worksheet, InputSheetModel inputSheetModel)
    {
        var dayDateFormat = "MMMM dd";

        var days = DayOfTheWeekColumnDictionary.Select((day, ix) =>
            new CellToAdd
            {
                Column = day.Value,
                Row = TitleRow,
                Value = $"{inputSheetModel.StartDate.AddDays(ix).ToString(dayDateFormat)}",
                Color = StyleConstants.DarkGrey
            }).ToList();

        CellToAdd[] cells =
        [
            new CellToAdd { Column = NameColInt, Row = TitleRow, Value = OutputWorkbookConstants.EmployeeNameTitle, Color = StyleConstants.DarkGrey},
            new CellToAdd { Column = HoursTypeColInt, Row = TitleRow, Value = OutputWorkbookConstants.HoursTypeTitle, Color = StyleConstants.DarkGrey, },
            ..days,
            new CellToAdd
            {
                Column = TotalHoursColInt, Row = TitleRow, Value = OutputWorkbookConstants.EmployeeTotalHoursTitle, Color = StyleConstants.DarkGrey,
            },
            new CellToAdd { Column = RatesColInt, Row = TitleRow, Value = OutputWorkbookConstants.EmployeeRatesTitle, Color = StyleConstants.DarkGrey, },
            new CellToAdd
            {
                Column = TotalBillColInt, Row = TitleRow, Value = OutputWorkbookConstants.EmployeeTotalBillTitle, Color = StyleConstants.DarkGrey,
            },
        ];

        foreach (var cell in cells)
        {
            worksheet.Cell(cell.Row, cell.Column).Value = cell.Value;
            worksheet.Cell(cell.Row, cell.Column).Style.Fill.BackgroundColor = cell.Color;
        }
    }

    private void AddDayRow(IXLWorksheet worksheet)
    {
        foreach (var day in Enum.GetValues(typeof(DayOfTheWeek)).Cast<DayOfTheWeek>())
        {
            worksheet.Cell(DayRow, DayOfTheWeekColumnDictionary[day]).Style.Fill.BackgroundColor =
                StyleConstants.LightGrey;
            worksheet.Cell(DayRow, DayOfTheWeekColumnDictionary[day]).Value = day.StringValue().ToUpper();
        }
    }

    private void AddWeekRow(IXLWorksheet worksheet, InputSheetModel inputSheetModel)
    {
        var weekRangeDateFormat = "dd MMMM yyyy";
        var formattedDate =
            $"{inputSheetModel.StartDate.ToString(weekRangeDateFormat)} - {inputSheetModel.EndDate.ToString(weekRangeDateFormat)}";
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
            worksheet.Cell(employeeRow, HoursTypeColInt).Value = OutputWorkbookConstants.PayrollHours;
            worksheet.Cell(employeeRow + 1, HoursTypeColInt).Value = OutputWorkbookConstants.RegularHours;
            worksheet.Cell(employeeRow + 2, HoursTypeColInt).Value = OutputWorkbookConstants.WeeklyOtHours;

            foreach (var (dayOfTheWeek, col) in DayOfTheWeekColumnDictionary)
            {
                worksheet.Cell(employeeRow, col).FormulaA1 =
                    $"{col.ToExcelColumn()}{employeeRow + 1} + {col.ToExcelColumn()}{employeeRow + 2}";
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
            worksheet.Cell(employeeRow + 1, TotalBillColInt).Style.NumberFormat.Format =
                StyleConstants.CurrencyCellFormat;

            worksheet.Cell(employeeRow + 2, TotalBillColInt).FormulaA1 =
                $"{TotalHoursColLetter}{employeeRow + 2} * {RatesColLetter}{employeeRow + 2}";
            ;
            worksheet.Cell(employeeRow + 2, TotalBillColInt).Style.NumberFormat.Format =
                StyleConstants.CurrencyCellFormat;
            employeeRow += 3;
        }
    }

    public void AddTotals(IXLWorksheet worksheet, InputSheetModel inputSheetModel, int totalsStartRow,
        int lastEmployeeRow)
    {
        var totalBillableRange = $"{TotalBillColLetter}1:{TotalBillColLetter}{lastEmployeeRow}";
        var mondayColLetter = DayOfTheWeekColumnDictionary[DayOfTheWeek.Monday].ToExcelColumn();
        var sundayColLetter = DayOfTheWeekColumnDictionary[DayOfTheWeek.Sunday].ToExcelColumn();
        var days = DayOfTheWeekColumnDictionary.SelectMany(ent =>
        {
            var col = ent.Value;
            var colLetter = col.ToExcelColumn();
            var daysHourRangeRange = $"{colLetter}1:{colLetter}{lastEmployeeRow}";

            return new[]
            {
                new CellToAdd
                {
                    Row = totalsStartRow + 1,
                    Column = col,
                    Value = $"{colLetter}{totalsStartRow + 2} + {colLetter}{totalsStartRow + 3}",
                    ValueType = CellToAdd.CellValueType.Formula,
                    NumberFormat = StyleConstants.HourCellFormat
                },
                new CellToAdd
                {
                    Row = totalsStartRow + 2,
                    Column = col,
                    Value = GetSumRangeBasedOnHourType(daysHourRangeRange,
                        OutputWorkbookConstants.RegularHours, lastEmployeeRow),
                    ValueType = CellToAdd.CellValueType.Formula,
                    NumberFormat = StyleConstants.HourCellFormat
                },
                new CellToAdd
                {
                    Row = totalsStartRow + 3,
                    Column = col,
                    Value = GetSumRangeBasedOnHourType(daysHourRangeRange,
                        OutputWorkbookConstants.WeeklyOtHours, lastEmployeeRow),
                    ValueType = CellToAdd.CellValueType.Formula,
                    NumberFormat = StyleConstants.HourCellFormat
                },
            };
        });

        CellToAdd[] cells =
        [
            new CellToAdd
            {
                Row = totalsStartRow, Column = HoursTypeColInt, Value = OutputWorkbookConstants.TotalHours
            },
            new CellToAdd { Row = totalsStartRow, Column = RatesColInt, Value = OutputWorkbookConstants.TotalBillable },
            new CellToAdd
            {
                Row = totalsStartRow,
                Column = TotalBillColInt,
                NumberFormat = StyleConstants.CurrencyWithSymbolCellFormat,
                Value = GetSumRangeBasedOnHourType(totalBillableRange, OutputWorkbookConstants.PayrollHours,
                    lastEmployeeRow),
                ValueType = CellToAdd.CellValueType.Formula
            },
            new CellToAdd { Row = totalsStartRow, Column = RatesColInt, Value = OutputWorkbookConstants.TotalBillable },
            new CellToAdd
            {
                Row = totalsStartRow + 1,
                Column = HoursTypeColInt,
                Value = OutputWorkbookConstants.TotalPayrollHours
            },
            new CellToAdd
            {
                Row = totalsStartRow + 2,
                Column = HoursTypeColInt,
                Value = OutputWorkbookConstants.TotalRegularHours
            },
            new CellToAdd
            {
                Row = totalsStartRow + 3,
                Column = HoursTypeColInt,
                Value = OutputWorkbookConstants.TotalWeeklyOtHours
            },
            ..days,
            new CellToAdd
            {
                Row = totalsStartRow + 1,
                Column = TotalHoursColInt,
                Value = $"SUM({mondayColLetter}{totalsStartRow + 1}:{sundayColLetter}{totalsStartRow + 1})",
                ValueType = CellToAdd.CellValueType.Formula,
                NumberFormat = StyleConstants.HourCellFormat
            },
            new CellToAdd
            {
                Row = totalsStartRow + 2,
                Column = TotalHoursColInt,
                Value = $"SUM({mondayColLetter}{totalsStartRow + 2}:{sundayColLetter}{totalsStartRow + 2})",
                ValueType = CellToAdd.CellValueType.Formula,
                NumberFormat = StyleConstants.HourCellFormat
            },new CellToAdd
            {
                Row = totalsStartRow + 3,
                Column = TotalHoursColInt,
                Value = $"SUM({mondayColLetter}{totalsStartRow + 3}:{sundayColLetter}{totalsStartRow + 3})",
                ValueType = CellToAdd.CellValueType.Formula,
                NumberFormat = StyleConstants.HourCellFormat
            },
        ];

        worksheet.Row(totalsStartRow).Style.Font.Bold = true;
        worksheet.Row(totalsStartRow).Style.Font.FontSize = StyleConstants.LargeFontSize;

        foreach (var cell in cells)
        {
            if (cell.ValueType == CellToAdd.CellValueType.Formula)
            {
                worksheet.Cell(cell.Row, cell.Column).FormulaA1 = cell.Value;
            }
            else
            {
                worksheet.Cell(cell.Row, cell.Column).Value = cell.Value;
            }

            worksheet.Cell(cell.Row, cell.Column).Style.NumberFormat.Format = cell.NumberFormat;
        }
    }

    private string GetSumRangeBasedOnHourType(string dataCellRange, string hourType, int lastEmployeeRow)
    {
        var typeRange = $"{HoursTypeColLetter}1:{HoursTypeColLetter}{lastEmployeeRow}";
        return $"SUMIFS({dataCellRange},{typeRange},\"{hourType}\")";
    }
}

public interface IWorksheetWriterHelper
{
    void AddTitleRows(IXLWorksheet worksheet, InputSheetModel inputSheetModel);
    void AddEmployeeRows(IXLWorksheet worksheet, IEnumerable<Employee> employees, ref int employeeRow);
    void AddTotals(IXLWorksheet worksheet, InputSheetModel inputSheetModel, int totalsStartRow, int lastEmployeeRow);
}
