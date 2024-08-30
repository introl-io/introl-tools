namespace Introl.Tools.Racks.Constants;

public class RackResultConstants
{
    public static int LabelColumn => 1;
    public static int IconColumn => 6;
    public static PortConstants SourcePort => new PortConstants
    {
        RackColumn = 2,
        RackUnitColumn = 3,
        SwitchColumn = 4,
        PortColumn = 5
    };
    public static PortConstants DestinationPort => new PortConstants
    {
        RackColumn = 7,
        RackUnitColumn = 8,
        SwitchColumn = 9,
        PortColumn = 10
    };

    public class PortConstants
    {
        public required int RackColumn { get; init; }
        public required int RackUnitColumn { get; init; }
        public required int SwitchColumn { get; init; }
        public required int PortColumn { get; init; }
    }
}
