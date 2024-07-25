using ClosedXML.Excel;
using Introl.Timesheets.Api.Constants;
using Introl.Timesheets.Api.Enums;
using Introl.Timesheets.Api.Extensions;
using Introl.Timesheets.Api.Models;
using Introl.Timesheets.Api.Timesheets.Team.Models;

namespace Introl.Timesheets.Api.Timesheets.Team.Services;

public class TeamResultCellFactory : ITeamResultCellFactory
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

    public IEnumerable<CellToAdd> GetTitleCells(IXLWorksheet worksheet, TeamParsedSourceModel teamSourceModel)
    {
        var pic = worksheet.AddPicture("./Assets/introl_logo.png")
            .MoveTo(1, 1);
        pic.Width = DimensionConstants.ImageWightAndHeightInPixels;
        pic.Height = DimensionConstants.ImageWightAndHeightInPixels;

        var cells = new List<CellToAdd>();
        for (var i = NameColInt; i <= TotalBillColInt; i++)
        {
            cells.Add(new CellToAdd
            {
                Column = i,
                Row = BufferRow,
                Color = StyleConstants.Black
            });
        }


        return
        [
            .. cells,
            .. AddWeekRow(teamSourceModel),
            .. AddDayRow(),
            .. AddTitleRow(teamSourceModel)
        ];
    }

    private IEnumerable<CellToAdd> AddTitleRow(TeamParsedSourceModel teamSourceModel)
    {
        var dayDateFormat = "MMMM dd";

        var days = DayOfTheWeekColumnDictionary.Select((day, ix) =>
            new CellToAdd
            {
                Column = day.Value,
                Row = TitleRow,
                Value = $"{teamSourceModel.StartDate.AddDays(ix).ToString(dayDateFormat)}",
                Color = StyleConstants.DarkGrey,
                Bold = true
            }).ToList();

        CellToAdd[] cells =
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
            .. days,
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
        ];

        return cells;
    }

    private IEnumerable<CellToAdd> AddDayRow()
    {
        foreach (var day in Enum.GetValues(typeof(DayOfTheWeek)).Cast<DayOfTheWeek>())
        {
            yield return new CellToAdd
            {
                Column = DayOfTheWeekColumnDictionary[day],
                Row = DayRow,
                Bold = true,
                Value = day.StringValue().ToUpper(),
                Color = StyleConstants.LightGrey
            };
        }
    }

    private IEnumerable<CellToAdd> AddWeekRow(TeamParsedSourceModel teamSourceModel)
    {
        var weekRangeDateFormat = "dd MMMM yyyy";
        var formattedDate =
            $"{teamSourceModel.StartDate.ToString(weekRangeDateFormat)} - {teamSourceModel.EndDate.ToString(weekRangeDateFormat)}";

        return new[] { new CellToAdd { Column = 2, Row = WeekRow, Bold = true, Value = formattedDate } };
    }

    public IEnumerable<CellToAdd> GetEmployeeCells(IEnumerable<TeamEmployee> employees, ref int employeeRow)
    {
        var mondayColLetter = DayOfTheWeekColumnDictionary[DayOfTheWeek.Monday].ToExcelColumn();
        var sundayColLetter = DayOfTheWeekColumnDictionary[DayOfTheWeek.Sunday].ToExcelColumn();

        var cells = new List<CellToAdd>();

        foreach (var employee in employees)
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

            foreach (var (dayOfTheWeek, col) in DayOfTheWeekColumnDictionary)
            {
                cells.AddRange(new[]
                {
                    new CellToAdd
                    {
                        Row = employeeRow,
                        Column = col,
                        ValueType = CellToAdd.CellValueType.Formula,
                        Value =
                            $"{col.ToExcelColumn()}{employeeRow + 1} + {col.ToExcelColumn()}{employeeRow + 2}",
                        NumberFormat = StyleConstants.HourCellFormat
                    },
                    new CellToAdd
                    {
                        Row = employeeRow + 1,
                        Column = col,
                        Value = employee.WorkDays[dayOfTheWeek].RegularHours,
                        NumberFormat = StyleConstants.HourCellFormat
                    },
                    new CellToAdd
                    {
                        Row = employeeRow + 2,
                        Column = col,
                        Value = employee.WorkDays[dayOfTheWeek].OvertimeHours,
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
                    Value = $"SUM({mondayColLetter}{employeeRow}:{sundayColLetter}{employeeRow})",
                    ValueType = CellToAdd.CellValueType.Formula,
                    NumberFormat = StyleConstants.HourCellFormat
                },
                new CellToAdd
                {
                    Row = employeeRow + 1,
                    Column = TotalHoursColInt,
                    Value = $"SUM({mondayColLetter}{employeeRow + 1}:{sundayColLetter}{employeeRow + 1})",
                    ValueType = CellToAdd.CellValueType.Formula,
                    NumberFormat = StyleConstants.HourCellFormat
                },
                new CellToAdd
                {
                    Row = employeeRow + 2,
                    Column = TotalHoursColInt,
                    Value = $"SUM({mondayColLetter}{employeeRow + 2}:{sundayColLetter}{employeeRow + 2})",
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

    public IEnumerable<CellToAdd> GetTotalsCells(TeamParsedSourceModel teamSourceModel, int totalsStartRow,
        int lastEmployeeRow)
    {
        var totalBillableRange = $"{TotalBillColLetter}1:{TotalBillColLetter}{lastEmployeeRow}";
        var mondayColLetter = DayOfTheWeekColumnDictionary[DayOfTheWeek.Monday].ToExcelColumn();
        var sundayColLetter = DayOfTheWeekColumnDictionary[DayOfTheWeek.Sunday].ToExcelColumn();
        var dayCells = DayOfTheWeekColumnDictionary.SelectMany(ent =>
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
            .. dayCells,
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
            },
            new CellToAdd
            {
                Row = totalsStartRow + 3,
                Column = TotalHoursColInt,
                Value = $"SUM({mondayColLetter}{totalsStartRow + 3}:{sundayColLetter}{totalsStartRow + 3})",
                ValueType = CellToAdd.CellValueType.Formula,
                NumberFormat = StyleConstants.HourCellFormat
            },
        ];


        return cells;
    }

    private string GetSumRangeBasedOnHourType(string dataCellRange, string hourType, int lastEmployeeRow)
    {
        var typeRange = $"{HoursTypeColLetter}1:{HoursTypeColLetter}{lastEmployeeRow}";
        return $"SUMIFS({dataCellRange},{typeRange},\"{hourType}\")";
    }
}

public interface ITeamResultCellFactory
{
    IEnumerable<CellToAdd> GetTitleCells(IXLWorksheet worksheet, TeamParsedSourceModel teamSourceModel);
    IEnumerable<CellToAdd> GetEmployeeCells(IEnumerable<TeamEmployee> employees, ref int employeeRow);
    IEnumerable<CellToAdd> GetTotalsCells(TeamParsedSourceModel teamSourceModel, int totalsStartRow, int lastEmployeeRow);
}
