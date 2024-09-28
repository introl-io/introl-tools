namespace Introl.Tools.Racks.Models;

public class ProcessFileRequest
{
    public required IFormFile File { get; init; }
    public required string SourcePortLabelFormat { get; init; }
    public required string DestinationPortLabelFormat { get; init; }
    public required bool HasHeadingRow { get; init; }

}
