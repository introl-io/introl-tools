namespace Introl.Tools.Racks.Models;

public class RackSourceModel
{
    public required IList<string> SourcePortHeadings { get; init; }
    public required IList<string> DestinationPortHeadings { get; init; }
    public required IList<RacksSourcePortMappingModel> PortMappings { get; init; }
}
