using ClosedXML.Excel;
using Introl.Timesheets.Api.Constants;
using Introl.Timesheets.Api.Models;
using Introl.Timesheets.Api.Timesheets.ActivityCode.Constants;
using Introl.Timesheets.Api.Timesheets.ActivityCode.Models;
using Introl.Timesheets.Api.Utils;

namespace Introl.Timesheets.Api.Timesheets.ActivityCode.Services;

public class ActCodeResultCellFactory(IActCodeHoursProcessor actCodeHoursProcessor) : IActCodeResultCellFactory
{
    public IEnumerable<CellToAdd> GetTitleCells(IXLWorksheet worksheet, ActCodeParsedSourceModel sourceModel,
        ref int row)
    {
        var pic = worksheet.AddPicture("./Assets/introl_logo.png")
            .MoveTo(ActCodeResultConstants.NameColInt, row);
        pic.Width = DimensionConstants.ImageWightAndHeightInPixels;
        pic.Height = DimensionConstants.ImageWightAndHeightInPixels;

        row++;

        row++;
        return
        [
            .. AddWeekRow(sourceModel, ref row),
            .. AddDayRow(sourceModel, ref row),
            .. AddTitleRow(sourceModel, ref row)
        ];
    }

    private IEnumerable<CellToAdd> AddTitleRow(ActCodeParsedSourceModel sourceModel, ref int row)
    {
        var dayDateFormat = "MMMM dd";

        var results = new List<CellToAdd>();
        var date = sourceModel.StartDate;
        var dayColumn = ActCodeResultConstants.DateStartColInt;

        while (date <= sourceModel.EndDate)
        {
            results.Add(new CellToAdd
            {
                Column = dayColumn,
                Row = row,
                Value = date.ToString(dayDateFormat),
                Color = StyleConstants.DarkGrey,
                Bold = true
            });
            dayColumn++;
            date = date.AddDays(1);
        }

        dayColumn--;

        results.AddRange(
        [
            new CellToAdd
            {
                Column = ActCodeResultConstants.NameColInt,
                Row = row,
                Value = OutputWorkbookConstants.EmployeeNameTitle,
                Color = StyleConstants.DarkGrey,
                Bold = true
            },
            new CellToAdd
            {
                Column = ActCodeResultConstants.EmployeeCodeInt,
                Row = row,
                Value = "Employee Code",
                Color = StyleConstants.DarkGrey,
                Bold = true
            },
            new CellToAdd
            {
                Column = ActCodeResultConstants.ActivityCodeColInt,
                Row = row,
                Value = "Activity Code",
                Color = StyleConstants.DarkGrey,
                Bold = true
            },
            new CellToAdd
            {
                Column = ActCodeResultConstants.HoursTypeColInt,
                Row = row,
                Value = OutputWorkbookConstants.HoursTypeTitle,
                Color = StyleConstants.DarkGrey,
                Bold = true
            },
            new CellToAdd
            {
                Column = dayColumn + ActCodeResultConstants.TotalHoursColOffset,
                Row = row,
                Value = OutputWorkbookConstants.EmployeeTotalHoursTitle,
                Color = StyleConstants.DarkGrey,
                Bold = true
            },
            new CellToAdd
            {
                Column = dayColumn + ActCodeResultConstants.RateColOffset,
                Row = row,
                Value = OutputWorkbookConstants.EmployeeRatesTitle,
                Color = StyleConstants.DarkGrey,
                Bold = true
            },
            new CellToAdd
            {
                Column = dayColumn + ActCodeResultConstants.TotalBillableColOffset,
                Row = row,
                Value = OutputWorkbookConstants.EmployeeTotalBillTitle,
                Color = StyleConstants.DarkGrey,
                Bold = true
            },
        ]);

        row++;
        return results;
    }

    private IEnumerable<CellToAdd> AddDayRow(ActCodeParsedSourceModel sourceModel, ref int row)
    {
        var col = ActCodeResultConstants.DateStartColInt;
        var date = sourceModel.StartDate;
        var results = new List<CellToAdd>();
        while (date <= sourceModel.EndDate)
        {
            results.Add(new CellToAdd
            {
                Column = col,
                Row = row,
                Bold = true,
                Value = date.ToString("ddd"),
                Color = StyleConstants.LightGrey
            });
            col++;
            date = date.AddDays(1);
        }

        row++;
        return results;
    }

    private IEnumerable<CellToAdd> AddWeekRow(ActCodeParsedSourceModel sourceModel, ref int row)
    {
        var weekRangeDateFormat = "dd MMMM yyyy";
        var formattedDate =
            $"{sourceModel.StartDate.ToString(weekRangeDateFormat)} - {sourceModel.EndDate.ToString(weekRangeDateFormat)}";

        var result = new[]
        {
            new CellToAdd
            {
                Column = ActCodeResultConstants.NameColInt, Row = row, Bold = true, Value = formattedDate
            }
        };
        row++;
        return result;
    }

    public IEnumerable<CellToAdd> CreateEmployeeCells(ActCodeParsedSourceModel sourceModel, ref int row)
    {
        var cells = new List<CellToAdd>();

        foreach (var employee in sourceModel.Employees)
        {
            cells.AddRange(CreateEmployeeCells(employee.Value, sourceModel.StartDate, sourceModel.EndDate,
                ref row));
        }

        return cells;
    }

    private IEnumerable<CellToAdd> CreateEmployeeCells(ActCodeEmployee employee, DateOnly startDate,
        DateOnly endDate,
        ref int row)
    {
        var processedHoursDict = actCodeHoursProcessor.Process(employee.ActivityCodeHours);

        var result = new List<CellToAdd>();
        result.AddRange(
        [
            new CellToAdd { Row = row, Column = ActCodeResultConstants.NameColInt, Value = employee.Name, Bold = true },
            new CellToAdd { Row = row, Column = ActCodeResultConstants.EmployeeCodeInt, Value = employee.MemberCode, Bold = true }
        ]);

        row++;
        foreach (var (activityCode, hoursDictionary) in processedHoursDict)
        {
            result.AddRange(CreateActCodeCells(activityCode, hoursDictionary, employee.Rate, startDate, endDate,
                ref row));
        }

        return result;
    }

    private IEnumerable<CellToAdd> CreateActCodeCells(string activityCode,
        Dictionary<DateOnly, (double regHours, double otHours)> actCodeHours, double rate, DateOnly startDate,
        DateOnly endDate,
        ref int row)
    {
        var cells = new List<CellToAdd>();

        cells.Add(new CellToAdd
        {
            Row = row,
            Column = ActCodeResultConstants.ActivityCodeColInt,
            Value = activityCode,
            Bold = true
        });
        row++;

        cells.Add(new CellToAdd
        {
            Row = row + ActCodeResultConstants.PayrollHoursOffset,
            Column = ActCodeResultConstants.HoursTypeColInt,
            Value = "Payroll Hours",
            Bold = true
        });

        cells.Add(new CellToAdd
        {
            Row = row + ActCodeResultConstants.OtHoursOffset,
            Column = ActCodeResultConstants.HoursTypeColInt,
            Value = "OT Hours",
            Bold = true
        });

        cells.Add(new CellToAdd
        {
            Row = row + ActCodeResultConstants.TotalHoursOffset,
            Column = ActCodeResultConstants.HoursTypeColInt,
            Value = "Total Hours",
            Bold = true
        });

        var date = startDate;
        var column = ActCodeResultConstants.DateStartColInt;
        while (date <= endDate)
        {
            var regHours = 0d;
            var otHours = 0d;
            if (actCodeHours.TryGetValue(date, out var hours))
            {
                regHours = hours.regHours;
                otHours = hours.otHours;
            }

            cells.Add(new CellToAdd
            {
                Row = row + ActCodeResultConstants.PayrollHoursOffset,
                Column = column,
                Value = regHours,
                NumberFormat = StyleConstants.HourCellFormat
            });
            cells.Add(new CellToAdd
            {
                Row = row + ActCodeResultConstants.OtHoursOffset,
                Column = column,
                Value = otHours,
                NumberFormat = StyleConstants.HourCellFormat
            });


            cells.Add(new CellToAdd
            {
                Row = row + ActCodeResultConstants.TotalHoursOffset,
                Column = column,
                ValueType = CellToAdd.CellValueType.Formula,
                NumberFormat = StyleConstants.HourCellFormat,
                Value =
                    $"=SUM({ExcelUtils.GetCellLocation(column, row)},{ExcelUtils.GetCellLocation(column, row + 1)})"
            });
            date = date.AddDays(1);
            column++;
        }

        cells.AddRange([
            new CellToAdd
            {
                Row = row + ActCodeResultConstants.PayrollHoursOffset,
                Column = column,
                ValueType = CellToAdd.CellValueType.Formula,
                NumberFormat = StyleConstants.HourCellFormat,
                Value =
                    $"=SUM({ExcelUtils.GetCellLocation(ActCodeResultConstants.DateStartColInt, row + ActCodeResultConstants.PayrollHoursOffset)}:{ExcelUtils.GetCellLocation(column - 1, row + ActCodeResultConstants.PayrollHoursOffset)})"
            },
            new CellToAdd
            {
                Row = row + ActCodeResultConstants.OtHoursOffset,
                Column = column,
                ValueType = CellToAdd.CellValueType.Formula,
                NumberFormat = StyleConstants.HourCellFormat,
                Value =
                    $"=SUM({ExcelUtils.GetCellLocation(ActCodeResultConstants.DateStartColInt, row + ActCodeResultConstants.TotalHoursColOffset)}:{ExcelUtils.GetCellLocation(column - 1, row + ActCodeResultConstants.OtHoursOffset)})"
            },
            new CellToAdd
            {
                Row = row + ActCodeResultConstants.TotalHoursOffset,
                Column = column,
                ValueType = CellToAdd.CellValueType.Formula,
                NumberFormat = StyleConstants.HourCellFormat,
                Value =
                    $"=SUM({ExcelUtils.GetCellLocation(ActCodeResultConstants.DateStartColInt, row + ActCodeResultConstants.TotalHoursOffset)}:{ExcelUtils.GetCellLocation(column - 1, row + ActCodeResultConstants.TotalHoursOffset)})"
            }
        ]);
        cells.AddRange(AddRatesCells(rate, column + 1, row));
        cells.AddRange(AddTotalBillableCells(column + 2, row));
        row += ActCodeResultConstants.ActCodeTotalRows;
        return cells;
    }

    private IEnumerable<CellToAdd> AddRatesCells(double rate, int column, int row)
    {
        return
        [
            new CellToAdd
            {
                Row = row + ActCodeResultConstants.PayrollHoursOffset,
                Column = column,
                Value = rate,
                NumberFormat = StyleConstants.CurrencyCellFormat
            },
            new CellToAdd
            {
                Row = row + ActCodeResultConstants.OtHoursOffset,
                Column = column,
                ValueType = CellToAdd.CellValueType.Formula,
                Value = $"{ExcelUtils.GetCellLocation(column, row + ActCodeResultConstants.PayrollHoursOffset)} * 1.5",
                NumberFormat = StyleConstants.CurrencyCellFormat
            }
        ];
    }

    private IEnumerable<CellToAdd> AddTotalBillableCells(int column, int row)
    {
        return
        [
            new CellToAdd
            {
                Row = row + ActCodeResultConstants.PayrollHoursOffset,
                Column = column,
                Value =
                    $"{ExcelUtils.GetCellLocation(column - 1, row + ActCodeResultConstants.PayrollHoursOffset)} * {ExcelUtils.GetCellLocation(column - 2, row + ActCodeResultConstants.PayrollHoursOffset)}",
                ValueType = CellToAdd.CellValueType.Formula,
                NumberFormat = StyleConstants.CurrencyCellFormat
            },
            new CellToAdd
            {
                Row = row + ActCodeResultConstants.OtHoursOffset,
                Column = column,
                Value =
                    $"{ExcelUtils.GetCellLocation(column - 1, row + ActCodeResultConstants.OtHoursOffset)} * {ExcelUtils.GetCellLocation(column - 2, row + ActCodeResultConstants.OtHoursOffset)}",
                ValueType = CellToAdd.CellValueType.Formula,
                NumberFormat = StyleConstants.CurrencyCellFormat
            },
            new CellToAdd
            {
                Row = row + ActCodeResultConstants.TotalHoursOffset,
                Column = column,
                Value =
                    $"{ExcelUtils.GetCellLocation(column, row + ActCodeResultConstants.PayrollHoursOffset)} + {ExcelUtils.GetCellLocation(column, row + ActCodeResultConstants.OtHoursOffset)}",
                ValueType = CellToAdd.CellValueType.Formula,
                NumberFormat = StyleConstants.CurrencyCellFormat
            }
        ];
    }
}

public interface IActCodeResultCellFactory
{
    IEnumerable<CellToAdd> GetTitleCells(IXLWorksheet worksheet, ActCodeParsedSourceModel sourceModel, ref int row);
    IEnumerable<CellToAdd> CreateEmployeeCells(ActCodeParsedSourceModel sourceModel, ref int row);
}
