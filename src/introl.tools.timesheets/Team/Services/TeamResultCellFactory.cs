using ClosedXML.Excel;
using Introl.Tools.Common.Constants;
using Introl.Tools.Common.Models;
using Introl.Tools.Common.Utils;
using Introl.Tools.Timesheets.Constants;
using Introl.Tools.Timesheets.Team.Models;

namespace Introl.Tools.Timesheets.Team.Services;

public class TeamResultCellFactory : ITeamResultCellFactory
{
    private const int NameColInt = 1;
    private const int HoursTypeColInt = 2;
    private const int DateStartColumn = 3;
    private const int TotalHoursColInt = 10;
    private const int RatesColInt = 11;
    private const int TotalBillColInt = 12;

    private string HoursTypeColLetter => ExcelUtils.ToExcelColumn(HoursTypeColInt);
    private string TotalHoursColLetter => ExcelUtils.ToExcelColumn(TotalHoursColInt);
    private string RatesColLetter => ExcelUtils.ToExcelColumn(RatesColInt);
    private string TotalBillColLetter => ExcelUtils.ToExcelColumn(TotalBillColInt);

    private const int BufferRow = 2;
    private const int WeekRow = 3;
    private const int DayRow = 4;
    private const int TitleRow = 5;

    public IEnumerable<CellToAdd> GetTitleCells(IXLWorksheet worksheet, TeamParsedSourceModel sourceModel)
    {
        var pic = worksheet.AddPicture("./Assets/introl_logo.png")
            .MoveTo(1, 1);
        pic.Width = DimensionConstants.ImageHeightInPixels;
        pic.Height = DimensionConstants.ImageWidthInPixels;

        var cells = new List<CellToAdd>();
        for (var i = NameColInt; i <= TotalBillColInt; i++)
        {
            cells.Add(new CellToAdd { Column = i, Row = BufferRow, Color = StyleConstants.Black });
        }


        return
        [
            .. cells,
            .. AddWeekRow(sourceModel),
            .. AddDayRow(sourceModel),
            .. AddTitleRow(sourceModel)
        ];
    }

    private IEnumerable<CellToAdd> AddTitleRow(TeamParsedSourceModel sourceModel)
    {
        var dayDateFormat = "MMMM dd";

        var numDays = GetNumberOfDays(sourceModel.StartDate, sourceModel.EndDate);
        var result = new List<CellToAdd>();
        for (var i = 0; i < numDays; i++)
        {
            result.Add(
                new CellToAdd
                {
                    Column = DateStartColumn + i,
                    Row = TitleRow,
                    Value = $"{sourceModel.StartDate.AddDays(i).ToString(dayDateFormat)}",
                    Color = StyleConstants.DarkGrey,
                    Bold = true
                });
        }


        result.AddRange(
        [
            new CellToAdd
            {
                Column = NameColInt,
                Row = TitleRow,
                Value = OutputWorkbookConstants.EmployeeNameTitle,
                Color = StyleConstants.DarkGrey,
                Bold = true
            },
            new CellToAdd
            {
                Column = HoursTypeColInt,
                Row = TitleRow,
                Value = OutputWorkbookConstants.HoursTypeTitle,
                Color = StyleConstants.DarkGrey,
                Bold = true
            },
            new CellToAdd
            {
                Column = TotalHoursColInt,
                Row = TitleRow,
                Value = OutputWorkbookConstants.EmployeeTotalHoursTitle,
                Color = StyleConstants.DarkGrey,
                Bold = true
            },
            new CellToAdd
            {
                Column = RatesColInt,
                Row = TitleRow,
                Value = OutputWorkbookConstants.EmployeeRatesTitle,
                Color = StyleConstants.DarkGrey,
                Bold = true
            },
            new CellToAdd
            {
                Column = TotalBillColInt,
                Row = TitleRow,
                Value = OutputWorkbookConstants.EmployeeTotalBillTitle,
                Color = StyleConstants.DarkGrey,
                Bold = true
            },
        ]);

        return result;
    }

    private IEnumerable<CellToAdd> AddDayRow(TeamParsedSourceModel sourceModel)
    {
        var numDays = GetNumberOfDays(sourceModel.StartDate, sourceModel.EndDate);
        var result = new List<CellToAdd>();
        for (var i = 0; i < numDays; i++)
        {
            result.Add(
                new CellToAdd
                {
                    Column = DateStartColumn + i,
                    Row = DayRow,
                    Bold = true,
                    Value = sourceModel.StartDate.AddDays(i).ToString("ddd").ToUpper(),
                    Color = StyleConstants.LightGrey
                }
            );
        }

        return result;
    }

    private IEnumerable<CellToAdd> AddWeekRow(TeamParsedSourceModel teamSourceModel)
    {
        var weekRangeDateFormat = "dd MMMM yyyy";
        var formattedDate =
            $"{teamSourceModel.StartDate.ToString(weekRangeDateFormat)} - {teamSourceModel.EndDate.ToString(weekRangeDateFormat)}";

        return new[] { new CellToAdd { Column = 2, Row = WeekRow, Bold = true, Value = formattedDate } };
    }

    public IEnumerable<CellToAdd> GetEmployeeCells(TeamParsedSourceModel sourceModel, ref int employeeRow)
    {
        var cells = new List<CellToAdd>();

        foreach (var employee in sourceModel.Employees)
        {
            cells.AddRange(new[]
                {
                    new CellToAdd
                    {
                        Value = employee.Name, Column = NameColInt, Row = employeeRow, Bold = true,
                    },
                    new
                        CellToAdd
                        {
                            Value = OutputWorkbookConstants.PayrollHours,
                            Column = HoursTypeColInt,
                            Row = employeeRow,
                            Bold = true,
                        },
                    new
                        CellToAdd
                        {
                            Value = OutputWorkbookConstants.RegularHours,
                            Column = HoursTypeColInt,
                            Row = employeeRow + 1,
                        },
                    new
                        CellToAdd
                        {
                            Value = OutputWorkbookConstants.WeeklyOtHours,
                            Column = HoursTypeColInt,
                            Row = employeeRow + 2,
                        }
                }
            );

            var numDays = GetNumberOfDays(sourceModel.StartDate, sourceModel.EndDate);

            for (var i = 0; i < numDays; i++)
            {
                var column = DateStartColumn + i;
                var date = sourceModel.StartDate.AddDays(i);
                cells.AddRange(new[]
                {
                    new CellToAdd
                    {
                        Row = employeeRow,
                        Column = column,
                        ValueType = CellToAdd.CellValueType.Formula,
                        Value =
                            $"{ExcelUtils.GetCellLocation(column, employeeRow + 1)} + {ExcelUtils.GetCellLocation(column, employeeRow + 2)}",
                        NumberFormat = StyleConstants.HourCellFormat
                    },
                    new CellToAdd
                    {
                        Row = employeeRow + 1,
                        Column = column,
                        Value = employee.WorkDays[date].RegularHours,
                        NumberFormat = StyleConstants.HourCellFormat
                    },
                    new CellToAdd
                    {
                        Row = employeeRow + 2,
                        Column = DateStartColumn + i,
                        Value = employee.WorkDays[date].OvertimeHours,
                        NumberFormat = StyleConstants.HourCellFormat
                    }
                });
            }

            cells.AddRange(new[]
            {
                new CellToAdd
                {
                    Row = employeeRow,
                    Column = TotalHoursColInt,
                    Value =
                        $"SUM({ExcelUtils.GetCellLocation(DateStartColumn, employeeRow)}:{ExcelUtils.GetCellLocation(DateStartColumn + numDays -1 , employeeRow)})",
                    ValueType = CellToAdd.CellValueType.Formula,
                    NumberFormat = StyleConstants.HourCellFormat
                },
                new CellToAdd
                {
                    Row = employeeRow + 1,
                    Column = TotalHoursColInt,
                    Value =
                        $"SUM({ExcelUtils.GetCellLocation(DateStartColumn, employeeRow + 1)}:{ExcelUtils.GetCellLocation(DateStartColumn + numDays -1, employeeRow + 1)})",
                    ValueType = CellToAdd.CellValueType.Formula,
                    NumberFormat = StyleConstants.HourCellFormat
                },
                new CellToAdd
                {
                    Row = employeeRow + 2,
                    Column = TotalHoursColInt,
                    Value =
                        $"SUM({ExcelUtils.GetCellLocation(DateStartColumn, employeeRow + 2)}:{ExcelUtils.GetCellLocation(DateStartColumn + numDays - 1, employeeRow + 2)})",
                    ValueType = CellToAdd.CellValueType.Formula,
                    NumberFormat = StyleConstants.HourCellFormat
                },
                new CellToAdd
                {
                    Row = employeeRow + 1,
                    Column = RatesColInt,
                    Value = employee.RegularHoursRate,
                    NumberFormat = StyleConstants.CurrencyCellFormat
                },
                new CellToAdd
                {
                    Row = employeeRow + 2,
                    Column = RatesColInt,
                    Value = employee.OvertimeHoursRate,
                    NumberFormat = StyleConstants.CurrencyCellFormat
                },
                new CellToAdd
                {
                    Row = employeeRow,
                    Column = TotalBillColInt,
                    Value = $"{TotalBillColLetter}{employeeRow + 1} + {TotalBillColLetter}{employeeRow + 2}",
                    ValueType = CellToAdd.CellValueType.Formula,
                    NumberFormat = StyleConstants.CurrencyCellFormat
                },
                new CellToAdd
                {
                    Row = employeeRow + 1,
                    Column = TotalBillColInt,
                    Value = $"{TotalHoursColLetter}{employeeRow + 1} * {RatesColLetter}{employeeRow + 1}",
                    ValueType = CellToAdd.CellValueType.Formula,
                    NumberFormat = StyleConstants.CurrencyCellFormat
                },
                new CellToAdd
                {
                    Row = employeeRow + 2,
                    Column = TotalBillColInt,
                    Value = $"{TotalHoursColLetter}{employeeRow + 2} * {RatesColLetter}{employeeRow + 2}",
                    ValueType = CellToAdd.CellValueType.Formula,
                    NumberFormat = StyleConstants.CurrencyCellFormat
                }
            });

            employeeRow += 3;
        }

        return cells;
    }

    public IEnumerable<CellToAdd> GetTotalsCells(TeamParsedSourceModel sourceModel, int totalsStartRow,
        int lastEmployeeRow)
    {
        var totalBillableRange = $"{TotalBillColLetter}1:{TotalBillColLetter}{lastEmployeeRow}";
        var result = new List<CellToAdd>();

        var numDays = GetNumberOfDays(sourceModel.StartDate, sourceModel.EndDate);
        var startDateColLetter = ExcelUtils.ToExcelColumn(DateStartColumn);
        var endDateColLetter = ExcelUtils.ToExcelColumn(DateStartColumn + numDays - 1);
        for (var i = 0; i < numDays; i++)
        {
            var col = DateStartColumn + i;
            var colLetter = ExcelUtils.ToExcelColumn(col);
            var daysHourRangeRange = $"{ExcelUtils.GetCellLocation(col, 1)}:{ExcelUtils.GetCellLocation(col, lastEmployeeRow)}";

            result.AddRange([
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
            ]);
        }

        result.AddRange(
        [
            new CellToAdd
            {
                Row = totalsStartRow,
                Column = HoursTypeColInt,
                Value = OutputWorkbookConstants.TotalHours,
                Bold = true,
                FontSize = StyleConstants.LargeFontSize
            },
            new CellToAdd
            {
                Row = totalsStartRow,
                Column = RatesColInt,
                Value = OutputWorkbookConstants.TotalBillable,
                Bold = true,
                FontSize = StyleConstants.LargeFontSize
            },
            new CellToAdd
            {
                Row = totalsStartRow,
                Column = TotalBillColInt,
                NumberFormat = StyleConstants.CurrencyWithSymbolCellFormat,
                Value = GetSumRangeBasedOnHourType(totalBillableRange, OutputWorkbookConstants.PayrollHours,
                    lastEmployeeRow),
                ValueType = CellToAdd.CellValueType.Formula,
                Bold = true,
                FontSize = StyleConstants.LargeFontSize
            },
            new CellToAdd
            {
                Row = totalsStartRow,
                Column = RatesColInt,
                Value = OutputWorkbookConstants.TotalBillable,
                Bold = true,
                FontSize = StyleConstants.LargeFontSize
            },
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
            new CellToAdd
            {
                Row = totalsStartRow + 1,
                Column = TotalHoursColInt,
                Value = $"SUM({startDateColLetter}{totalsStartRow + 1}:{endDateColLetter}{totalsStartRow + 1})",
                ValueType = CellToAdd.CellValueType.Formula,
                NumberFormat = StyleConstants.HourCellFormat
            },
            new CellToAdd
            {
                Row = totalsStartRow + 2,
                Column = TotalHoursColInt,
                Value = $"SUM({startDateColLetter}{totalsStartRow + 2}:{endDateColLetter}{totalsStartRow + 2})",
                ValueType = CellToAdd.CellValueType.Formula,
                NumberFormat = StyleConstants.HourCellFormat
            },
            new CellToAdd
            {
                Row = totalsStartRow + 3,
                Column = TotalHoursColInt,
                Value = $"SUM({startDateColLetter}{totalsStartRow + 3}:{endDateColLetter}{totalsStartRow + 3})",
                ValueType = CellToAdd.CellValueType.Formula,
                NumberFormat = StyleConstants.HourCellFormat
            },
        ]);


        return result;
    }

    private string GetSumRangeBasedOnHourType(string dataCellRange, string hourType, int lastEmployeeRow)
    {
        var typeRange = $"{HoursTypeColLetter}1:{HoursTypeColLetter}{lastEmployeeRow}";
        return $"SUMIFS({dataCellRange},{typeRange},\"{hourType}\")";
    }

    private int GetNumberOfDays(DateOnly startDate, DateOnly endDate)
    {
        return endDate.DayNumber - startDate.DayNumber + 1;
    }
}

public interface ITeamResultCellFactory
{
    IEnumerable<CellToAdd> GetTitleCells(IXLWorksheet worksheet, TeamParsedSourceModel sourceModel);
    IEnumerable<CellToAdd> GetEmployeeCells(TeamParsedSourceModel sourceModel, ref int employeeRow);

    IEnumerable<CellToAdd> GetTotalsCells(TeamParsedSourceModel sourceModel, int totalsStartRow,
        int lastEmployeeRow);
}
