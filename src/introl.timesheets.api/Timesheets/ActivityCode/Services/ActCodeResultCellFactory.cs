using Introl.Timesheets.Api.Models;
using Introl.Timesheets.Api.Timesheets.ActivityCode.Models;
using Introl.Timesheets.Api.Utils;

namespace Introl.Timesheets.Api.Timesheets.ActivityCode.Services;

public class ActCodeResultCellFactory(IActCodeHoursProcessor actCodeHoursProcessor) : IActCodeResultCellFactory
{
    public IEnumerable<CellToAdd> CreateEmployeeCells(ActCodeParsedSourceModel sourceModel, ref int row)
    {
        var cells = new List<CellToAdd>();

        foreach (var employee in sourceModel.Employees)
        {
            cells.AddRange(CreateEmployeeCells(employee.Value, sourceModel.StartDate, sourceModel.EndDate, ref row));
        }

        return cells;
    }

    private IEnumerable<CellToAdd> CreateEmployeeCells(ActCodeEmployee employee, DateOnly startDate, DateOnly endDate,
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
        var date = startDate;

        cells.Add(new CellToAdd { Row = row, Column = 2, Value = activityCode, Bold = true });
        row++;

        cells.Add(new CellToAdd { Row = row, Column = 3, Value = "Payroll Hours", Bold = true });

        cells.Add(new CellToAdd { Row = row + 1, Column = 3, Value = "OT Hours", Bold = true });

        cells.Add(new CellToAdd { Row = row + 2, Column = 3, Value = "Total Hours", Bold = true });

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
                Value = $"=SUM({ExcelUtils.GetCellLocation(4, row)}:{ExcelUtils.GetCellLocation(column - 1, row)})"
            },
            new CellToAdd
            {
                Row = row+1,
                Column = column,
                ValueType = CellToAdd.CellValueType.Formula,
                Value = $"=SUM({ExcelUtils.GetCellLocation(4, row+1)}:{ExcelUtils.GetCellLocation(column - 1, row+1)})"
            },
            new CellToAdd
            {
                Row = row+2,
                Column = column,
                ValueType = CellToAdd.CellValueType.Formula,
                Value = $"=SUM({ExcelUtils.GetCellLocation(4, row+2)}:{ExcelUtils.GetCellLocation(column - 1, row+2)})"
            }
        ]);

        row += 3;
        return cells;
    }
}

public interface IActCodeResultCellFactory
{
    IEnumerable<CellToAdd> CreateEmployeeCells(ActCodeParsedSourceModel sourceModel, ref int row);
}
