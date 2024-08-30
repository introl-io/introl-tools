namespace Introl.Tools.Racks.Models;

public sealed class RacksSourcePortMappingModel
{
    public required PortModel SourcePort { get; init; }
    public required PortModel DestinationPort { get; init; }

    public sealed class PortModel
    {
        public required string Rack { get; init; }
        public required string RackUnit { get; init; }
        public required string Switch { get; init; }
        public required string Port { get; init; }
    }
}
