using Introl.Timesheets.Api.Enums;

namespace Introl.Timesheets.Api.Models;

public class ProcessedTimesheetError
{
    public required TimesheetProcessingFailureReasons FailureReason { get; init; }
    public required string Message { get; init; }
}
