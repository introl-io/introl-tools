namespace Introl.Timesheets.Api.Models;

public class CellToAdd
{
    public required int Column { get; init; }
    public required int Row { get; init; }
    public required string Value { get; init; }
    public bool Bold = false;
}
