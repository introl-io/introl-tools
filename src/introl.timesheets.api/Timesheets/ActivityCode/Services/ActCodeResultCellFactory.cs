using ClosedXML.Excel;
using Introl.Timesheets.Api.Constants;
using Introl.Timesheets.Api.Models;
using Introl.Timesheets.Api.Timesheets.ActivityCode.Models;
using Introl.Timesheets.Api.Utils;

namespace Introl.Timesheets.Api.Timesheets.ActivityCode.Services;

public class ActCodeResultCellFactory(IActCodeHoursProcessor actCodeHoursProcessor) : IActCodeResultCellFactory
{
    public IEnumerable<CellToAdd> GetTitleCells(IXLWorksheet worksheet, ActCodeParsedSourceModel sourceModel,
        ref int row)
    {
        var pic = worksheet.AddPicture("./Assets/introl_logo.png")
            .MoveTo(1, row);
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
        var column = 4;
        
        while (date <= sourceModel.EndDate)
        {
            results.Add(new CellToAdd
            {
                Column = column,
                Row = row,
                Value = date.ToString(dayDateFormat),
                Color = StyleConstants.DarkGrey,
                Bold = true
            });
            column++;
            date = date.AddDays(1);
        }


        results.AddRange(
            [
                new CellToAdd
                {
                    Column = 1,
                    Row = row,
                    Value = OutputWorkbookConstants.EmployeeNameTitle,
                    Color = StyleConstants.DarkGrey,
                    Bold = true
                },new CellToAdd
                {
                    Column = 2,
                    Row = row,
                    Value = "Activity Code",
                    Color = StyleConstants.DarkGrey,
                    Bold = true
                },
                new CellToAdd
                {
                    Column = 3,
                    Row = row,
                    Value = OutputWorkbookConstants.HoursTypeTitle,
                    Color = StyleConstants.DarkGrey,
                    Bold = true
                },
                new CellToAdd
                {
                    Column = column,
                    Row = row,
                    Value = OutputWorkbookConstants.EmployeeTotalHoursTitle,
                    Color = StyleConstants.DarkGrey,
                    Bold = true
                },
                new CellToAdd
                {
                    Column = column+1,
                    Row = row,
                    Value = OutputWorkbookConstants.EmployeeRatesTitle,
                    Color = StyleConstants.DarkGrey,
                    Bold = true
                },
                new CellToAdd
                {
                    Column = column+2,
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
            var col = 4;
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

            var result = new[] { new CellToAdd { Column = 2, Row = row, Bold = true, Value = formattedDate } };
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
            result.Add(new CellToAdd { Row = row, Column = 1, Value = employee.Name, Bold = true });

            row++;
            foreach (var (activityCode, hoursDictionary) in processedHoursDict)
            {
                result.AddRange(CreateActCodeCells(activityCode, hoursDictionary, startDate, endDate, ref row));
            }

            return result;
        }

        private IEnumerable<CellToAdd> CreateActCodeCells(string activityCode,
            Dictionary<DateOnly, (double regHours, double otHours)> actCodeHours, DateOnly startDate, DateOnly endDate,
            ref int row)
        {
            var cells = new List<CellToAdd>();

            cells.Add(new CellToAdd { Row = row, Column = 2, Value = activityCode, Bold = true });
            row++;

            cells.Add(new CellToAdd { Row = row, Column = 3, Value = "Payroll Hours", Bold = true });

            cells.Add(new CellToAdd { Row = row + 1, Column = 3, Value = "OT Hours", Bold = true });

            cells.Add(new CellToAdd { Row = row + 2, Column = 3, Value = "Total Hours", Bold = true });

            var date = startDate;
            var column = 4;
            while (date <= endDate)
            {
                if (actCodeHours.TryGetValue(date, out var hours))
                {
                    cells.Add(new CellToAdd { Row = row, Column = column, Value = hours.regHours });
                    cells.Add(new CellToAdd { Row = row + 1, Column = column, Value = hours.otHours });
                }
                else
                {
                    cells.Add(new CellToAdd { Row = row, Column = column, Value = 0 });
                    cells.Add(new CellToAdd { Row = row + 1, Column = column, Value = 0 });
                }

                cells.Add(new CellToAdd
                {
                    Row = row + 2,
                    Column = column,
                    ValueType = CellToAdd.CellValueType.Formula,
                    Value =
                        $"=SUM({ExcelUtils.GetCellLocation(column, row)},{ExcelUtils.GetCellLocation(column, row + 1)})"
                });
                date = date.AddDays(1);
                column++;
            }

            cells.AddRange([
                new CellToAdd
                {
                    Row = row,
                    Column = column,
                    ValueType = CellToAdd.CellValueType.Formula,
                    Value =
                        $"=SUM({ExcelUtils.GetCellLocation(4, row)}:{ExcelUtils.GetCellLocation(column - 1, row)})"
                },
                new CellToAdd
                {
                    Row = row + 1,
                    Column = column,
                    ValueType = CellToAdd.CellValueType.Formula,
                    Value =
                        $"=SUM({ExcelUtils.GetCellLocation(4, row + 1)}:{ExcelUtils.GetCellLocation(column - 1, row + 1)})"
                },
                new CellToAdd
                {
                    Row = row + 2,
                    Column = column,
                    ValueType = CellToAdd.CellValueType.Formula,
                    Value =
                        $"=SUM({ExcelUtils.GetCellLocation(4, row + 2)}:{ExcelUtils.GetCellLocation(column - 1, row + 2)})"
                }
            ]);

            row += 3;
            return cells;
        }
    }

    public interface IActCodeResultCellFactory
    {
        IEnumerable<CellToAdd> GetTitleCells(IXLWorksheet worksheet, ActCodeParsedSourceModel sourceModel, ref int row);
        IEnumerable<CellToAdd> CreateEmployeeCells(ActCodeParsedSourceModel sourceModel, ref int row);
    }
