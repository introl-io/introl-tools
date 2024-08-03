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

        var hoursTypeFirstCell = ExcelUtils.GetCellLocation(ActCodeResultConstants.HoursTypeColInt, employeeFirstRow);
        var hoursTypeFinalCell = ExcelUtils.GetCellLocation(ActCodeResultConstants.HoursTypeColInt, employeeFinalRow);

        var actCodeFirstCell = ExcelUtils.GetCellLocation(ActCodeResultConstants.ActivityCodeColInt, employeeFirstRow);
        var actCodeFinalCell = ExcelUtils.GetCellLocation(ActCodeResultConstants.ActivityCodeColInt, employeeFinalRow);

        var titleColumn = ActCodeResultConstants.DateStartColInt - 2;
        cells.Add(new CellToAdd
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
            cells.Add(new CellToAdd { Column = titleColumn, Row = row, Value = actCode, Bold = true });
            row++;
            cells.Add(new CellToAdd
            {
                Column = ActCodeResultConstants.HoursTypeColInt,
                Row = row + ActCodeResultConstants.RegularHoursOffset,
                Value = OutputWorkbookConstants.RegularHours,
            });

            cells.Add(new CellToAdd
            {
                Column = ActCodeResultConstants.HoursTypeColInt,
                Row = row + ActCodeResultConstants.OtHoursOffset,
                Value = OutputWorkbookConstants.WeeklyOtHours,
            });

            cells.Add(new CellToAdd
            {
                Column = ActCodeResultConstants.HoursTypeColInt,
                Row = row + ActCodeResultConstants.PayrollHoursOffset,
                Value = OutputWorkbookConstants.PayrollHours,
                Bold = true
            });
            var regHoursCell = ExcelUtils.GetCellLocation(ActCodeResultConstants.HoursTypeColInt, row + ActCodeResultConstants.RegularHoursOffset);
            var otHoursCell = ExcelUtils.GetCellLocation(ActCodeResultConstants.HoursTypeColInt, row + ActCodeResultConstants.OtHoursOffset);

            var date = sourceModel.StartDate;
            var column = ActCodeResultConstants.DateStartColInt;
            while (date <= sourceModel.EndDate)
            {
                var employeeFirstCell = ExcelUtils.GetCellLocation(column, employeeFirstRow);
                var employeeFinalCell = ExcelUtils.GetCellLocation(column, employeeFinalRow);
                cells.Add(new CellToAdd
                {
                    Row = row + ActCodeResultConstants.RegularHoursOffset,
                    Column = column,
                    ValueType = CellToAdd.CellValueType.Formula,
                    Value =
                        $"SUMIFS({employeeFirstCell}:{employeeFinalCell}, {hoursTypeFirstCell}:{hoursTypeFinalCell}, {regHoursCell}, {actCodeFirstCell}:{actCodeFinalCell}, {actCodeCell})",
                    NumberFormat = StyleConstants.HourCellFormat
                });
                cells.Add(new CellToAdd
                {
                    Row = row + ActCodeResultConstants.OtHoursOffset,
                    Column = column,
                    ValueType = CellToAdd.CellValueType.Formula,
                    Value =
                        $"SUMIFS({employeeFirstCell}:{employeeFinalCell}, {hoursTypeFirstCell}:{hoursTypeFinalCell}, {otHoursCell}, {actCodeFirstCell}:{actCodeFinalCell}, {actCodeCell})",
                    NumberFormat = StyleConstants.HourCellFormat
                });

                cells.Add(new CellToAdd
                {
                    Row = row + ActCodeResultConstants.PayrollHoursOffset,
                    Column = column,
                    ValueType = CellToAdd.CellValueType.Formula,
                    Value =
                        $"{ExcelUtils.GetCellLocation(column, row + ActCodeResultConstants.RegularHoursOffset)} + {ExcelUtils.GetCellLocation(column, row + ActCodeResultConstants.OtHoursOffset)}",
                    NumberFormat = StyleConstants.HourCellFormat
                });
                date = date.AddDays(1);
                column++;
            }

            cells.Add(new CellToAdd
            {
                Row = row + ActCodeResultConstants.RegularHoursOffset,
                Column = column,
                ValueType = CellToAdd.CellValueType.Formula,
                Value =
                    $"SUM({ExcelUtils.GetCellLocation(ActCodeResultConstants.DateStartColInt, row + ActCodeResultConstants.RegularHoursOffset)}:{ExcelUtils.GetCellLocation(column - 1, row + ActCodeResultConstants.RegularHoursOffset)})",
                NumberFormat = StyleConstants.HourCellFormat
            });
            cells.Add(new CellToAdd
            {
                Row = row + ActCodeResultConstants.OtHoursOffset,
                Column = column,
                ValueType = CellToAdd.CellValueType.Formula,
                Value =
                    $"SUM({ExcelUtils.GetCellLocation(ActCodeResultConstants.DateStartColInt, row + ActCodeResultConstants.OtHoursOffset)}:{ExcelUtils.GetCellLocation(column - 1, row + ActCodeResultConstants.OtHoursOffset)})",
                NumberFormat = StyleConstants.HourCellFormat
            });
            cells.Add(new CellToAdd
            {
                Row = row + ActCodeResultConstants.PayrollHoursOffset,
                Column = column,
                ValueType = CellToAdd.CellValueType.Formula,
                Value =
                    $"SUM({ExcelUtils.GetCellLocation(ActCodeResultConstants.DateStartColInt, row + ActCodeResultConstants.PayrollHoursOffset)}:{ExcelUtils.GetCellLocation(column - 1, row + ActCodeResultConstants.PayrollHoursOffset)})",
                NumberFormat = StyleConstants.HourCellFormat
            });

            row += ActCodeResultConstants.TotalBlockTotalRows;
        }

        var finalActCodeRow = row;
        cells.Add(new CellToAdd { Column = titleColumn, Row = row, Value = "Total", Bold = true });
        row++;
        cells.AddRange([
            new CellToAdd
            {
                Column = ActCodeResultConstants.HoursTypeColInt,
                Row = row + ActCodeResultConstants.RegularHoursOffset,
                Value = OutputWorkbookConstants.RegularHours,
                Bold = true
            },
            new CellToAdd
            {
                Column = ActCodeResultConstants.HoursTypeColInt,
                Row = row + ActCodeResultConstants.OtHoursOffset,
                Value = OutputWorkbookConstants.WeeklyOtHours,
                Bold = true
            },
            new CellToAdd
            {
                Column = ActCodeResultConstants.HoursTypeColInt,
                Row = row + ActCodeResultConstants.PayrollHoursOffset,
                Value = OutputWorkbookConstants.PayrollHours,
                Bold = true
            },
        ]);
        
        var date2 = sourceModel.StartDate;
        var column2 = ActCodeResultConstants.DateStartColInt;
        
        var regHoursCell2 = ExcelUtils.GetCellLocation(ActCodeResultConstants.HoursTypeColInt, row + ActCodeResultConstants.RegularHoursOffset);
        var otHoursCell2 = ExcelUtils.GetCellLocation(ActCodeResultConstants.HoursTypeColInt, row + ActCodeResultConstants.OtHoursOffset);
        var payrollHoursCell2 = ExcelUtils.GetCellLocation(ActCodeResultConstants.HoursTypeColInt, row + ActCodeResultConstants.PayrollHoursOffset);
        while (date2 <= sourceModel.EndDate)
        {
            var firstHoursCell = ExcelUtils.GetCellLocation(column2, employeeFinalRow + 1);
            var finalHoursCell = ExcelUtils.GetCellLocation(column2, finalActCodeRow);
            
            var firstActCodeCell = ExcelUtils.GetCellLocation(ActCodeResultConstants.HoursTypeColInt, employeeFinalRow + 1);
            var finalActCodeCell = ExcelUtils.GetCellLocation(ActCodeResultConstants.HoursTypeColInt, finalActCodeRow);
            cells.AddRange([
            new CellToAdd
            {
                Row = row + ActCodeResultConstants.RegularHoursOffset,
                Column = column2,
                ValueType = CellToAdd.CellValueType.Formula,
                Value = $"SUMIFS({firstHoursCell}:{finalHoursCell}, {firstActCodeCell}:{finalActCodeCell}, {regHoursCell2})"
            },
            new CellToAdd
            {
                Row = row + ActCodeResultConstants.OtHoursOffset,
                Column = column2,
                ValueType = CellToAdd.CellValueType.Formula,
                Value = $"SUMIFS({firstHoursCell}:{finalHoursCell}, {firstActCodeCell}:{finalActCodeCell}, {otHoursCell2})"
            },
            new CellToAdd
            {
                Row = row + ActCodeResultConstants.PayrollHoursOffset,
                Column = column2,
                ValueType = CellToAdd.CellValueType.Formula,
                Value = $"SUMIFS({firstHoursCell}:{finalHoursCell}, {firstActCodeCell}:{finalActCodeCell}, {payrollHoursCell2})"
            }
            ]);
            column2++;
            date2 = date2.AddDays(1);
        }
        
        return cells;
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
        var weekRangeDateFormat = "dd MMMM yyyy";
        var formattedDate =
            $"{sourceModel.StartDate.ToString(weekRangeDateFormat)} - {sourceModel.EndDate.ToString(weekRangeDateFormat)}";

        var col = ActCodeResultConstants.DateStartColInt;
        var date = sourceModel.StartDate;
        var results = new List<CellToAdd>();

        results.Add(
            new CellToAdd
            {
                Column = ActCodeResultConstants.NameColInt, Row = row, Bold = true, Value = formattedDate
            });
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

    private IEnumerable<CellToAdd> CreateEmployeeCells(ActCodeEmployee employee, DateOnly startDate,
        DateOnly endDate,
        ref int row)
    {
        var processedHoursDict = actCodeHoursProcessor.Process(employee.ActivityCodeHours);

        var result = new List<CellToAdd>();
        result.AddRange(
        [
            new CellToAdd { Row = row, Column = ActCodeResultConstants.NameColInt, Value = employee.Name, Bold = true },
            new CellToAdd
            {
                Row = row, Column = ActCodeResultConstants.EmployeeCodeInt, Value = employee.MemberCode, Bold = true
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
                NumberFormat = i == row ? null : StyleConstants.HideTextCellFormat,
                Bold = true
            });
        }

        row++;

        cells.Add(new CellToAdd
        {
            Row = row + ActCodeResultConstants.RegularHoursOffset,
            Column = ActCodeResultConstants.HoursTypeColInt,
            Value = OutputWorkbookConstants.RegularHours,
        });

        cells.Add(new CellToAdd
        {
            Row = row + ActCodeResultConstants.OtHoursOffset,
            Column = ActCodeResultConstants.HoursTypeColInt,
            Value = OutputWorkbookConstants.WeeklyOtHours,
        });

        cells.Add(new CellToAdd
        {
            Row = row + ActCodeResultConstants.PayrollHoursOffset,
            Column = ActCodeResultConstants.HoursTypeColInt,
            Value = OutputWorkbookConstants.PayrollHours,
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
                Row = row + ActCodeResultConstants.RegularHoursOffset,
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
                Row = row + ActCodeResultConstants.PayrollHoursOffset,
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
}

public interface IActCodeResultCellFactory
{
    IEnumerable<CellToAdd> GetTitleCells(IXLWorksheet worksheet, ActCodeParsedSourceModel sourceModel, ref int row);
    IEnumerable<CellToAdd> CreateEmployeeCells(ActCodeParsedSourceModel sourceModel, ref int row);

    IEnumerable<CellToAdd> CreateTotalsCell(ActCodeParsedSourceModel sourceModel, int employeeFirstRow,
        int employeeFinalRow, ref int row);
}
