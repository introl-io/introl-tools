using Introl.Timesheets.Api.Models;
using Introl.Timesheets.Api.Timesheets.ActivityCode.Models;

namespace Introl.Timesheets.Api.Timesheets.ActivityCode.Services;

public class ActCodeResultCellFactory :IActCodeResultCellFactory
{
    public IEnumerable<CellToAdd> CreateEmployeeCells(IEnumerable<ActCodeEmployee> employees)
    {
        throw new NotImplementedException();
    }
}

public interface IActCodeResultCellFactory
{
    IEnumerable<CellToAdd> CreateEmployeeCells(IEnumerable<ActCodeEmployee> employees);
}