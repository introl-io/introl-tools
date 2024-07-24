namespace Introl.Timesheets.Api.Models.EmployeeTimesheets;

public class ProcessedTimesheetResult
{
    public required string Name { get; init; }
    public required byte[] WorkbookBytes { get; init; }

}
