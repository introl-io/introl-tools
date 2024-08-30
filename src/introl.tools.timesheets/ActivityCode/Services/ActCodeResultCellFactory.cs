using ClosedXML.Excel;
using Introl.Tools.Common.Constants;
using Introl.Tools.Common.Models;
using Introl.Tools.Common.Utils;
using Introl.Tools.Timesheets.ActivityCode.Constants;
using Introl.Tools.Timesheets.ActivityCode.Models;
using Introl.Tools.Timesheets.Constants;

namespace Introl.Tools.Timesheets.ActivityCode.Services;

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

        return
        [
            .. AddDayRow(sourceModel, ref row),
            .. AddTitleRow(sourceModel, ref row)
        ];
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

    public IEnumerable<CellToAdd> CreateTotalsCell(ActCodeParsedSourceModel sourceModel, int employeeFirstRow,
        int employeeFinalRow, ref int row)
    {
        var cells = new List<CellToAdd>();

        var titleColumn = ActCodeResultConstants.DateStartColInt - 2;

        cells.AddRange(CreateTotalActCodeCells(sourceModel, employeeFirstRow, employeeFinalRow, titleColumn, ref row));

        var finalActCodeRow = row;
        cells.Add(new CellToAdd { Column = titleColumn, Row = row, Value = "Total", Bold = true });
        row++;
        cells.AddRange([
            new CellToAdd
            {
                Column = ActCodeResultConstants.HoursTypeColInt,
                Row = row + ActCodeResultConstants.RegularHoursOffset,
                Value = OutputWorkbookConstants.RegularHours,
            },
            new CellToAdd
            {
                Column = ActCodeResultConstants.HoursTypeColInt,
                Row = row + ActCodeResultConstants.OtHoursOffset,
                Value = OutputWorkbookConstants.WeeklyOtHours,
            },
            new CellToAdd
            {
                Column = ActCodeResultConstants.HoursTypeColInt,
                Row = row + ActCodeResultConstants.PayrollHoursOffset,
                Value = OutputWorkbookConstants.PayrollHours,
                Bold = true
            },
        ]);

        var column = ActCodeResultConstants.DateStartColInt;

        var regHoursCell = ExcelUtils.GetCellLocation(ActCodeResultConstants.HoursTypeColInt,
            row + ActCodeResultConstants.RegularHoursOffset);
        var otHoursCell = ExcelUtils.GetCellLocation(ActCodeResultConstants.HoursTypeColInt,
            row + ActCodeResultConstants.OtHoursOffset);
        var payrollHoursCell = ExcelUtils.GetCellLocation(ActCodeResultConstants.HoursTypeColInt,
            row + ActCodeResultConstants.PayrollHoursOffset);

        var firstActCodeCell =
            ExcelUtils.GetCellLocation(ActCodeResultConstants.HoursTypeColInt, employeeFinalRow + 1);
        var finalActCodeCell = ExcelUtils.GetCellLocation(ActCodeResultConstants.HoursTypeColInt, finalActCodeRow);

        var numDays = GetNumberOfDays(sourceModel.StartDate, sourceModel.EndDate);

        for (var i = 0; i < numDays; i++)
        {
            var firstHoursCell = ExcelUtils.GetCellLocation(column, employeeFinalRow + 1);
            var finalHoursCell = ExcelUtils.GetCellLocation(column, finalActCodeRow);
            cells.AddRange([
                new CellToAdd
                {
                    Row = row + ActCodeResultConstants.RegularHoursOffset,
                    Column = column,
                    ValueType = CellToAdd.CellValueType.Formula,
                    NumberFormat = StyleConstants.HourCellFormat,
                    Value =
                        $"SUMIFS({firstHoursCell}:{finalHoursCell}, {firstActCodeCell}:{finalActCodeCell}, {regHoursCell})"
                },
                new CellToAdd
                {
                    Row = row + ActCodeResultConstants.OtHoursOffset,
                    Column = column,
                    ValueType = CellToAdd.CellValueType.Formula,
                    NumberFormat = StyleConstants.HourCellFormat,
                    Value =
                        $"SUMIFS({firstHoursCell}:{finalHoursCell}, {firstActCodeCell}:{finalActCodeCell}, {otHoursCell})"
                },
                new CellToAdd
                {
                    Row = row + ActCodeResultConstants.PayrollHoursOffset,
                    Column = column,
                    ValueType = CellToAdd.CellValueType.Formula,
                    NumberFormat = StyleConstants.HourCellFormat,
                    Value =
                        $"SUMIFS({firstHoursCell}:{finalHoursCell}, {firstActCodeCell}:{finalActCodeCell}, {payrollHoursCell})"
                }
            ]);
            column++;
        }

        var totalBillableColumn = column + ActCodeResultConstants.TotalBillableColOffset - 1;
        var totalHoursFirstCell = ExcelUtils.GetCellLocation(column, employeeFinalRow + 1);
        var totalHoursFinalCell = ExcelUtils.GetCellLocation(column, finalActCodeRow);

        var totalBillableFirstCell = ExcelUtils.GetCellLocation(totalBillableColumn, employeeFinalRow + 1);
        var totalBillableFinalCell = ExcelUtils.GetCellLocation(totalBillableColumn, finalActCodeRow);
        cells.AddRange([
            new CellToAdd
            {
                Column = column,
                Row = row + ActCodeResultConstants.RegularHoursOffset,
                ValueType = CellToAdd.CellValueType.Formula,
                NumberFormat = StyleConstants.HourCellFormat,
                Value =
                    $"SUMIFS({totalHoursFirstCell}:{totalHoursFinalCell}, {firstActCodeCell}:{finalActCodeCell}, {regHoursCell})"
            },
            new CellToAdd
            {
                Column = column,
                Row = row + ActCodeResultConstants.OtHoursOffset,
                ValueType = CellToAdd.CellValueType.Formula,
                NumberFormat = StyleConstants.HourCellFormat,
                Value =
                    $"SUMIFS({totalHoursFirstCell}:{totalHoursFinalCell}, {firstActCodeCell}:{finalActCodeCell}, {otHoursCell})"
            },
            new CellToAdd
            {
                Column = column,
                Row = row + ActCodeResultConstants.PayrollHoursOffset,
                ValueType = CellToAdd.CellValueType.Formula,
                NumberFormat = StyleConstants.HourCellFormat,
                Value =
                    $"{ExcelUtils.GetCellLocation(column, row + ActCodeResultConstants.RegularHoursOffset)} + {ExcelUtils.GetCellLocation(column, row + ActCodeResultConstants.OtHoursOffset)}"
            },
            new CellToAdd
            {
                Column = totalBillableColumn,
                Row = row + ActCodeResultConstants.RegularHoursOffset,
                ValueType = CellToAdd.CellValueType.Formula,
                NumberFormat = StyleConstants.CurrencyCellFormat,
                Value =
                    $"SUMIFS({totalBillableFirstCell}:{totalBillableFinalCell}, {firstActCodeCell}:{finalActCodeCell}, {regHoursCell})"
            },
            new CellToAdd
            {
                Column = totalBillableColumn,
                Row = row + ActCodeResultConstants.OtHoursOffset,
                ValueType = CellToAdd.CellValueType.Formula,
                NumberFormat = StyleConstants.CurrencyCellFormat,
                Value =
                    $"SUMIFS({totalBillableFirstCell}:{totalBillableFinalCell}, {firstActCodeCell}:{finalActCodeCell}, {otHoursCell})"
            },
            new CellToAdd
            {
                Column = totalBillableColumn,
                Row = row + ActCodeResultConstants.PayrollHoursOffset,
                ValueType = CellToAdd.CellValueType.Formula,
                NumberFormat = StyleConstants.CurrencyWithSymbolCellFormat,
                Bold = true,
                FontSize = StyleConstants.LargeFontSize,
                Value =
                    $"{ExcelUtils.GetCellLocation(totalBillableColumn, row + ActCodeResultConstants.RegularHoursOffset)} + {ExcelUtils.GetCellLocation(totalBillableColumn, row + ActCodeResultConstants.OtHoursOffset)}"
            },
        ]);
        return cells;
    }

    private IEnumerable<CellToAdd> CreateTotalActCodeCells(ActCodeParsedSourceModel sourceModel, int employeeFirstRow,
        int employeeFinalRow, int titleColumn, ref int row)
    {
        var result = new List<CellToAdd>();
        var hoursTypeFirstCell = ExcelUtils.GetCellLocation(ActCodeResultConstants.HoursTypeColInt, employeeFirstRow);
        var hoursTypeFinalCell = ExcelUtils.GetCellLocation(ActCodeResultConstants.HoursTypeColInt, employeeFinalRow);

        var actCodeFirstCell = ExcelUtils.GetCellLocation(ActCodeResultConstants.ActivityCodeColInt, employeeFirstRow);
        var actCodeFinalCell = ExcelUtils.GetCellLocation(ActCodeResultConstants.ActivityCodeColInt, employeeFinalRow);

        result.Add(new CellToAdd
        {
            Bold = true,
            Column = titleColumn,
            Row = row,
            Value = "Total Hours",
            FontSize = StyleConstants.LargeFontSize
        });
        row += 2;
        foreach (var actCode in sourceModel.ActivityCodes)
        {
            var actCodeCell = ExcelUtils.GetCellLocation(titleColumn, row);
            result.Add(new CellToAdd { Column = titleColumn, Row = row, Value = actCode, Bold = true });
            row++;
            result.AddRange(
            [
                new CellToAdd
                {
                    Column = ActCodeResultConstants.HoursTypeColInt,
                    Row = row + ActCodeResultConstants.RegularHoursOffset,
                    Value = OutputWorkbookConstants.RegularHours,
                },
                new CellToAdd
                {
                    Column = ActCodeResultConstants.HoursTypeColInt,
                    Row = row + ActCodeResultConstants.OtHoursOffset,
                    Value = OutputWorkbookConstants.WeeklyOtHours,
                },
                new CellToAdd
                {
                    Column = ActCodeResultConstants.HoursTypeColInt,
                    Row = row + ActCodeResultConstants.PayrollHoursOffset,
                    Value = OutputWorkbookConstants.PayrollHours,
                    Bold = true
                }
            ]);
            var regHoursCell = ExcelUtils.GetCellLocation(ActCodeResultConstants.HoursTypeColInt,
                row + ActCodeResultConstants.RegularHoursOffset);
            var otHoursCell = ExcelUtils.GetCellLocation(ActCodeResultConstants.HoursTypeColInt,
                row + ActCodeResultConstants.OtHoursOffset);

            var column = ActCodeResultConstants.DateStartColInt;

            var numDays = GetNumberOfDays(sourceModel.StartDate, sourceModel.EndDate);
            for (var i = 0; i < numDays; i++)
            {
                var employeeFirstCell = ExcelUtils.GetCellLocation(column, employeeFirstRow);
                var employeeFinalCell = ExcelUtils.GetCellLocation(column, employeeFinalRow);
                result.AddRange(
                [
                    new CellToAdd
                    {
                        Row = row + ActCodeResultConstants.RegularHoursOffset,
                        Column = column,
                        ValueType = CellToAdd.CellValueType.Formula,
                        Value =
                            $"SUMIFS({employeeFirstCell}:{employeeFinalCell}, {hoursTypeFirstCell}:{hoursTypeFinalCell}, {regHoursCell}, {actCodeFirstCell}:{actCodeFinalCell}, {actCodeCell})",
                        NumberFormat = StyleConstants.HourCellFormat
                    },
                    new CellToAdd
                    {
                        Row = row + ActCodeResultConstants.OtHoursOffset,
                        Column = column,
                        ValueType = CellToAdd.CellValueType.Formula,
                        Value =
                            $"SUMIFS({employeeFirstCell}:{employeeFinalCell}, {hoursTypeFirstCell}:{hoursTypeFinalCell}, {otHoursCell}, {actCodeFirstCell}:{actCodeFinalCell}, {actCodeCell})",
                        NumberFormat = StyleConstants.HourCellFormat
                    },
                    new CellToAdd
                    {
                        Row = row + ActCodeResultConstants.PayrollHoursOffset,
                        Column = column,
                        ValueType = CellToAdd.CellValueType.Formula,
                        Value =
                            $"{ExcelUtils.GetCellLocation(column, row + ActCodeResultConstants.RegularHoursOffset)} + {ExcelUtils.GetCellLocation(column, row + ActCodeResultConstants.OtHoursOffset)}",
                        NumberFormat = StyleConstants.HourCellFormat
                    }
                ]);
                column++;
            }

            var totalBillableCol = column + ActCodeResultConstants.TotalBillableColOffset - 1;
            var totalBillableFirstCell = ExcelUtils.GetCellLocation(totalBillableCol, employeeFirstRow);
            var totalBillableFinalCell = ExcelUtils.GetCellLocation(totalBillableCol, employeeFinalRow);

            result.AddRange([
                new CellToAdd
                {
                    Row = row + ActCodeResultConstants.RegularHoursOffset,
                    Column = column,
                    ValueType = CellToAdd.CellValueType.Formula,
                    Value =
                        $"SUM({ExcelUtils.GetCellLocation(ActCodeResultConstants.DateStartColInt, row + ActCodeResultConstants.RegularHoursOffset)}:{ExcelUtils.GetCellLocation(column - 1, row + ActCodeResultConstants.RegularHoursOffset)})",
                    NumberFormat = StyleConstants.HourCellFormat
                },
                new CellToAdd
                {
                    Row = row + ActCodeResultConstants.OtHoursOffset,
                    Column = column,
                    ValueType = CellToAdd.CellValueType.Formula,
                    Value =
                        $"SUM({ExcelUtils.GetCellLocation(ActCodeResultConstants.DateStartColInt, row + ActCodeResultConstants.OtHoursOffset)}:{ExcelUtils.GetCellLocation(column - 1, row + ActCodeResultConstants.OtHoursOffset)})",
                    NumberFormat = StyleConstants.HourCellFormat
                },
                new CellToAdd
                {
                    Row = row + ActCodeResultConstants.PayrollHoursOffset,
                    Column = column,
                    ValueType = CellToAdd.CellValueType.Formula,
                    Value =
                        $"SUM({ExcelUtils.GetCellLocation(ActCodeResultConstants.DateStartColInt, row + ActCodeResultConstants.PayrollHoursOffset)}:{ExcelUtils.GetCellLocation(column - 1, row + ActCodeResultConstants.PayrollHoursOffset)})",
                    NumberFormat = StyleConstants.HourCellFormat
                },
                new CellToAdd
                {
                    Row = row + ActCodeResultConstants.RegularHoursOffset,
                    Column = totalBillableCol,
                    ValueType = CellToAdd.CellValueType.Formula,
                    Value =
                        $"SUMIFS({totalBillableFirstCell}:{totalBillableFinalCell}, {hoursTypeFirstCell}:{hoursTypeFinalCell}, {regHoursCell}, {actCodeFirstCell}:{actCodeFinalCell}, {actCodeCell})",
                    NumberFormat = StyleConstants.CurrencyCellFormat
                },
                new CellToAdd
                {
                    Row = row + ActCodeResultConstants.OtHoursOffset,
                    Column = totalBillableCol,
                    ValueType = CellToAdd.CellValueType.Formula,
                    Value =
                        $"SUMIFS({totalBillableFirstCell}:{totalBillableFinalCell}, {hoursTypeFirstCell}:{hoursTypeFinalCell}, {otHoursCell}, {actCodeFirstCell}:{actCodeFinalCell}, {actCodeCell})",
                    NumberFormat = StyleConstants.CurrencyCellFormat
                },
                new CellToAdd
                {
                    Row = row + ActCodeResultConstants.PayrollHoursOffset,
                    Column = totalBillableCol,
                    ValueType = CellToAdd.CellValueType.Formula,
                    Value =
                        $"{ExcelUtils.GetCellLocation(totalBillableCol, row + ActCodeResultConstants.RegularHoursOffset)} + {ExcelUtils.GetCellLocation(totalBillableCol, row + ActCodeResultConstants.OtHoursOffset)}",
                    NumberFormat = StyleConstants.CurrencyWithSymbolCellFormat,
                    Bold = true,
                    FontSize = StyleConstants.LargeFontSize
                }
            ]);
            row += ActCodeResultConstants.TotalBlockTotalRows;
        }

        return result;
    }


    private IEnumerable<CellToAdd> AddTitleRow(ActCodeParsedSourceModel sourceModel, ref int row)
    {
        var dayDateFormat = "MMMM dd";

        var results = new List<CellToAdd>();
        var dayColumn = ActCodeResultConstants.DateStartColInt;

        var numDays = GetNumberOfDays(sourceModel.StartDate, sourceModel.EndDate);
        for (var i = 0; i < numDays; i++)
        {
            results.Add(new CellToAdd
            {
                Column = dayColumn,
                Row = row,
                Value = sourceModel.StartDate.AddDays(i).ToString(dayDateFormat),
                Color = StyleConstants.DarkGrey,
                Bold = true
            });
            dayColumn++;
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
        var weekRangeDateFormat = "dd MMMM yyyy";
        var formattedDate =
            $"{sourceModel.StartDate.ToString(weekRangeDateFormat)} - {sourceModel.EndDate.ToString(weekRangeDateFormat)}";

        var col = ActCodeResultConstants.DateStartColInt;
        var results = new List<CellToAdd>();

        results.Add(
            new CellToAdd
            {
                Column = ActCodeResultConstants.NameColInt,
                Row = row,
                Bold = true,
                Value = formattedDate
            });

        var numDays = GetNumberOfDays(sourceModel.StartDate, sourceModel.EndDate);
        for (var i = 0; i < numDays; i++)
        {
            results.Add(new CellToAdd
            {
                Column = col,
                Row = row,
                Bold = true,
                Value = sourceModel.StartDate.AddDays(i).ToString("ddd"),
                Color = StyleConstants.LightGrey
            });
            col++;
        }

        row++;
        return results;
    }

    private IEnumerable<CellToAdd> CreateEmployeeCells(ActCodeEmployee employee, DateOnly startDate,
        DateOnly endDate,
        ref int row)
    {
        var processedHoursDict = actCodeHoursProcessor.Process(employee.ActivityCodeHours);

        var result = new List<CellToAdd>();

        var totalNumCols = ActCodeResultConstants.DateStartColInt + GetNumberOfDays(startDate, endDate) +
                           ActCodeResultConstants.TotalBillableColOffset;
        for (var i = ActCodeResultConstants.NameColInt; i < totalNumCols; i++)
        {
            result.Add(new CellToAdd
            {
                Row = row,
                Column = i,
                Color = StyleConstants.MutedBlue
            });
        }

        result.AddRange(
        [
            new CellToAdd
            {
                Row = row,
                Column = ActCodeResultConstants.NameColInt,
                Value = employee.Name,
                Bold = true
            },
            new CellToAdd
            {
                Row = row,
                Column = ActCodeResultConstants.EmployeeCodeInt,
                Value = employee.MemberCode,
                Bold = true
            }
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

        for (var i = row; i <= row + ActCodeResultConstants.ActCodeTotalRows; i++)
        {
            cells.Add(new CellToAdd
            {
                Row = i,
                Column = ActCodeResultConstants.ActivityCodeColInt,
                Value = activityCode,
                // Hides the value, needs to be there so SUMIFS elsewhere can work
                NumberFormat = i == row ? null : StyleConstants.HideTextCellFormat,
                Bold = true
            });
        }
        var totalNumCols = ActCodeResultConstants.DateStartColInt + GetNumberOfDays(startDate, endDate) +
                           ActCodeResultConstants.TotalBillableColOffset;

        for (var i = ActCodeResultConstants.ActivityCodeColInt + 1; i < totalNumCols; i++)
        {
            cells.Add(new CellToAdd
            {
                Row = row,
                Column = i,
                Color = StyleConstants.LightGrey
            });
        }

        row++;

        cells.AddRange([
            new CellToAdd
            {
                Row = row + ActCodeResultConstants.RegularHoursOffset,
                Column = ActCodeResultConstants.HoursTypeColInt,
                Value = OutputWorkbookConstants.RegularHours,
            },
            new CellToAdd
            {
                Row = row + ActCodeResultConstants.OtHoursOffset,
                Column = ActCodeResultConstants.HoursTypeColInt,
                Value = OutputWorkbookConstants.WeeklyOtHours,
            },
            new CellToAdd
            {
                Row = row + ActCodeResultConstants.PayrollHoursOffset,
                Column = ActCodeResultConstants.HoursTypeColInt,
                Value = OutputWorkbookConstants.PayrollHours,
                Bold = true
            }
        ]);

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

            cells.AddRange(
            [
                new CellToAdd
                {
                    Row = row + ActCodeResultConstants.RegularHoursOffset,
                    Column = column,
                    Value = regHours,
                    NumberFormat = StyleConstants.HourCellFormat
                },
                new CellToAdd
                {
                    Row = row + ActCodeResultConstants.OtHoursOffset,
                    Column = column,
                    Value = otHours,
                    NumberFormat = StyleConstants.HourCellFormat
                },
                new CellToAdd
                {
                    Row = row + ActCodeResultConstants.PayrollHoursOffset,
                    Column = column,
                    ValueType = CellToAdd.CellValueType.Formula,
                    NumberFormat = StyleConstants.HourCellFormat,
                    Value =
                        $"=SUM({ExcelUtils.GetCellLocation(column, row)},{ExcelUtils.GetCellLocation(column, row + 1)})"
                }
            ]);
            date = date.AddDays(1);
            column++;
        }

        cells.AddRange([
            new CellToAdd
            {
                Row = row + ActCodeResultConstants.RegularHoursOffset,
                Column = column,
                ValueType = CellToAdd.CellValueType.Formula,
                NumberFormat = StyleConstants.HourCellFormat,
                Value =
                    $"=SUM({ExcelUtils.GetCellLocation(ActCodeResultConstants.DateStartColInt, row + ActCodeResultConstants.RegularHoursOffset)}:{ExcelUtils.GetCellLocation(column - 1, row + ActCodeResultConstants.RegularHoursOffset)})"
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
                Row = row + ActCodeResultConstants.PayrollHoursOffset,
                Column = column,
                ValueType = CellToAdd.CellValueType.Formula,
                NumberFormat = StyleConstants.HourCellFormat,
                Value =
                    $"=SUM({ExcelUtils.GetCellLocation(ActCodeResultConstants.DateStartColInt, row + ActCodeResultConstants.PayrollHoursOffset)}:{ExcelUtils.GetCellLocation(column - 1, row + ActCodeResultConstants.PayrollHoursOffset)})"
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
                Row = row + ActCodeResultConstants.RegularHoursOffset,
                Column = column,
                Value = rate,
                NumberFormat = StyleConstants.CurrencyCellFormat
            },
            new CellToAdd
            {
                Row = row + ActCodeResultConstants.OtHoursOffset,
                Column = column,
                ValueType = CellToAdd.CellValueType.Formula,
                Value =
                    $"{ExcelUtils.GetCellLocation(column, row + ActCodeResultConstants.RegularHoursOffset)} * 1.5",
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
                Row = row + ActCodeResultConstants.RegularHoursOffset,
                Column = column,
                Value =
                    $"{ExcelUtils.GetCellLocation(column - 1, row + ActCodeResultConstants.RegularHoursOffset)} * {ExcelUtils.GetCellLocation(column - 2, row + ActCodeResultConstants.RegularHoursOffset)}",
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
                Row = row + ActCodeResultConstants.PayrollHoursOffset,
                Column = column,
                Value =
                    $"{ExcelUtils.GetCellLocation(column, row + ActCodeResultConstants.RegularHoursOffset)} + {ExcelUtils.GetCellLocation(column, row + ActCodeResultConstants.OtHoursOffset)}",
                ValueType = CellToAdd.CellValueType.Formula,
                NumberFormat = StyleConstants.CurrencyCellFormat
            }
        ];
    }

    private int GetNumberOfDays(DateOnly startDate, DateOnly endDate)
    {
        return endDate.DayNumber - startDate.DayNumber + 1;
    }
}

public interface IActCodeResultCellFactory
{
    IEnumerable<CellToAdd> GetTitleCells(IXLWorksheet worksheet, ActCodeParsedSourceModel sourceModel, ref int row);
    IEnumerable<CellToAdd> CreateEmployeeCells(ActCodeParsedSourceModel sourceModel, ref int row);

    IEnumerable<CellToAdd> CreateTotalsCell(ActCodeParsedSourceModel sourceModel, int employeeFirstRow,
        int employeeFinalRow, ref int row);
}
