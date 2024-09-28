namespace Introl.Tools.Racks.Models;

public class RacksSourceKeyPositions
{
    public required int StartRow { get; init; }
    public required int EndRow { get; init; }
    public required string[] SourceColumns { get; init; }
    public required string[] DestinationColumns { get; init; }
}
