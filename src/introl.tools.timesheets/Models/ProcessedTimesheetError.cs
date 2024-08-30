using Introl.Tools.Timesheets.Enums;

namespace Introl.Tools.Timesheets.Models;

public class ProcessedTimesheetError
{
    public required TimesheetProcessingFailureReasons FailureReason { get; init; }
    public required string Message { get; init; }
}
