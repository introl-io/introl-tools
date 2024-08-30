namespace Introl.Tools.Racks.Models;

public sealed class RacksSourceModel
{
    public required PortModel SourcePort { get; init; }
    public required PortModel DestinationPort { get; init; }
    
    public sealed class PortModel
    {
        public required string Rack { get; init; }
        public required string RackUnit { get; init; }
        public required int Slot { get; init; }
        public required int Port { get; init; }
    }
}
