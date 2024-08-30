using Introl.Tools.Common.Enums;

namespace Introl.Tools.Common.Models;

public class ProcessingError
{
    public required ProcessingFailureReasons FailureReason { get; init; }
    public required string Message { get; init; }
}
