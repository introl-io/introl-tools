using Introl.Tools.Common.Constants;
using Introl.Tools.Common.Models;
using Introl.Tools.Common.Utils;
using Introl.Tools.Racks.Constants;
using Introl.Tools.Racks.Models;

namespace Introl.Tools.Racks.Services;

public class RackCellFactory : IRackCellFactory
{
    public IList<CellToAdd> GetTitleCells()
    {
        var result = new List<CellToAdd>
        {
            new CellToAdd
            {
                Row = 1,
                Column = RackResultConstants.LabelColumn,
                Value = "Label",
                Color = StyleConstants.MutedBlue,
                Bold = true
            },
            new CellToAdd
            {
                Row = 1,
                Column = RackResultConstants.IconColumn,
                Value = "",
                Color = StyleConstants.MutedBlue,
                Bold = true
            }
        };

        result.AddRange(GetPortTitleCells(RackResultConstants.SourcePort));
        result.AddRange(GetPortTitleCells(RackResultConstants.DestinationPort));

        return result;
    }

    private IList<CellToAdd> GetPortTitleCells(RackResultConstants.PortConstants portConstants) =>
        new List<CellToAdd>
        {
            new CellToAdd
            {
                Row = 1,
                Column = portConstants.RackColumn,
                Value = "Rack",
                Color = StyleConstants.MutedBlue,
                Bold = true
            },
            new CellToAdd
            {
                Row = 1,
                Column = portConstants.RackUnitColumn,
                Value = "RU",
                Color = StyleConstants.MutedBlue,
                Bold = true
            },
            new CellToAdd
            {
                Row = 1,
                Column = portConstants.SwitchColumn,
                Value = "Switch Port",
                Color = StyleConstants.MutedBlue,
                Bold = true
            },
            new CellToAdd
            {
                Row = 1,
                Column = portConstants.PortColumn,
                Value = "Optic Port",
                Color = StyleConstants.MutedBlue,
                Bold = true
            }
        };
    public IList<CellToAdd> GetPortMappingCells(RackSourceModel sourceModel, int startRow)
    {
        var result = new List<CellToAdd>();

        for (var i = 0; i < sourceModel.PortMappings.Count; i++)
        {
            var row = i + startRow;
            var portMapping = sourceModel.PortMappings[i];

            result.Add(GetLabelCell(row));
            result.AddRange(GetPortCells(portMapping.SourcePort, RackResultConstants.SourcePort, row));
            result.AddRange(GetPortCells(portMapping.DestinationPort, RackResultConstants.DestinationPort, row));
            result.Add(new CellToAdd
            {
                Row = row,
                Column = RackResultConstants.IconColumn,
                Value = "<->",
                Color = StyleConstants.Blue
            });
        }

        return result;
    }

    private CellToAdd GetLabelCell(int row)
    {
        var srcRackAddress = ExcelUtils.GetCellLocation(RackResultConstants.SourcePort.RackColumn, row);
        var destRackAddress = ExcelUtils.GetCellLocation(RackResultConstants.DestinationPort.RackColumn, row);

        var srcRUAddress = ExcelUtils.GetCellLocation(RackResultConstants.SourcePort.RackUnitColumn, row);
        var destRUAddress = ExcelUtils.GetCellLocation(RackResultConstants.DestinationPort.RackUnitColumn, row);

        var srcSwitchAddress = ExcelUtils.GetCellLocation(RackResultConstants.SourcePort.SwitchColumn, row);
        var destSwitchAddress = ExcelUtils.GetCellLocation(RackResultConstants.DestinationPort.SwitchColumn, row);

        var srcPortAddress = ExcelUtils.GetCellLocation(RackResultConstants.SourcePort.PortColumn, row);
        var destPortAddress = ExcelUtils.GetCellLocation(RackResultConstants.DestinationPort.PortColumn, row);

        var srcFormattedPortAddress =
            GetFormattedPortAddress(srcRackAddress, srcRUAddress, srcSwitchAddress, srcPortAddress);
        var destFormattedPortAddress =
            GetFormattedPortAddress(destRackAddress, destRUAddress, destSwitchAddress, destPortAddress);
        var value = $"{srcFormattedPortAddress} & CHAR(10) & \" - \" & CHAR(10) & {destFormattedPortAddress}";

        return new CellToAdd
        {
            Row = row,
            Column = RackResultConstants.LabelColumn,
            ValueType = CellToAdd.CellValueType.Formula,
            Value = value,
            Wrap = true
        };
    }

    private string GetFormattedPortAddress(string rackAddress, string ruAddress, string switchAddress,
        string portAddress) => $"{rackAddress} & \".\" & {ruAddress} & \".\" & {switchAddress} & \".\" & {portAddress}";

    private IList<CellToAdd> GetPortCells(
        RacksSourcePortMappingModel.PortModel portModel,
        RackResultConstants.PortConstants portConstants,
        int row)
    {
        return new List<CellToAdd>
        {
            new CellToAdd { Row = row, Column = portConstants.RackColumn, Value = portModel.Rack },
            new CellToAdd { Row = row, Column = portConstants.RackUnitColumn, Value = portModel.RackUnit },
            new CellToAdd { Row = row, Column = portConstants.SwitchColumn, Value = portModel.Switch },
            new CellToAdd { Row = row, Column = portConstants.PortColumn, Value = portModel.Port },
        };
    }
}

public interface IRackCellFactory
{
    IList<CellToAdd> GetTitleCells();

    IList<CellToAdd> GetPortMappingCells(RackSourceModel sourceModel, int startRow);
}
