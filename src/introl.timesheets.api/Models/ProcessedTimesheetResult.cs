namespace Introl.Timesheets.Api.Models;

public class ProcessedTimesheetResult
{
    public required string Name { get; init; }
    public required byte[] WorkbookBytes { get; init; }

}
