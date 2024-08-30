namespace Introl.Tools.Racks.Models;

public class RacksSourceKeyPositions
{
    public required int StartRow { get; init; }
    public required int EndRow { get; init; }
    public required PortKeyPositions SourcePort { get; init; }
    public required PortKeyPositions DestinationPort { get; init; }

    public sealed class PortKeyPositions
    {
        public required int RackColumn { get; init; }
        public required int RackUnitColumn { get; init; }
        public required int SlotColumn { get; init; }
        public required int PortColumn { get; init; }
    }
}
