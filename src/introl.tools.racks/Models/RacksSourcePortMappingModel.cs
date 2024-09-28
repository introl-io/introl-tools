namespace Introl.Tools.Racks.Models;

public sealed class RacksSourcePortMappingModel
{
    public required string[] SourcePort { get; init; }
    public required string[] DestinationPort { get; init; }
}
