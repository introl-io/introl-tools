using System.Text.RegularExpressions;
using Introl.Tools.Common.Constants;
using Introl.Tools.Common.Models;
using Introl.Tools.Racks.Models;

namespace Introl.Tools.Racks.Services;

public class RackCellFactory : IRackCellFactory
{
    public IList<CellToAdd> GetHeaderCells(RackSourceModel sourceModel)
    {
        var iconColumn = 2 + sourceModel.PortMappings.First().SourcePort.Length;
        var result = new List<CellToAdd>
        {
            new CellToAdd
            {
                Row = 1,
                Column = 1,
                Value = "Label",
                Color = StyleConstants.MutedBlue,
                Bold = true
            },
            new CellToAdd
            {
                Row = 1,
                Column = 2 + sourceModel.PortMappings.First().SourcePort.Length,
                Value = "",
                Color = StyleConstants.MutedBlue,
                Bold = true
            }
        };

        result.AddRange(GetPortTitleCells(sourceModel.SourcePortHeadings, 2));
        result.AddRange(GetPortTitleCells(sourceModel.SourcePortHeadings, iconColumn + 1));

        return result;
    }

    private IList<CellToAdd> GetPortTitleCells(IList<string> portHeadings, int startColumn) =>
        portHeadings
            .Select((p, ix) =>
                new CellToAdd
                {
                    Row = 1,
                    Column = ix + startColumn,
                    Value = p,
                    Color =
                    StyleConstants.MutedBlue,
                    Bold = true
                })
            .ToList();

    public IList<CellToAdd> GetPortMappingCells(RackSourceModel sourceModel,
        int startRow,
        string sourcePortLabelFormat,
        string destinationPortLabelFormat)
    {
        var result = new List<CellToAdd>();

        for (var i = 0; i < sourceModel.PortMappings.Count; i++)
        {
            var row = i + startRow;
            var portMapping = sourceModel.PortMappings[i];

            result.Add(GetLabelCell(row, portMapping, sourcePortLabelFormat, destinationPortLabelFormat));
            result.AddRange(GetPortCells(portMapping.SourcePort, row, 2));
            result.AddRange(GetPortCells(portMapping.DestinationPort, row, 3 + portMapping.SourcePort.Length));
            result.Add(new CellToAdd
            {
                Row = row,
                Column = 2 + portMapping.SourcePort.Length,
                Value = "<->",
                Color = StyleConstants.Blue
            });
        }

        return result;
    }

    private CellToAdd GetLabelCell(int row,
        RacksSourcePortMappingModel portMapping,
        string sourcePortLabelFormat,
        string destinationPortLabelFormat)
    {
        var srcFormattedPortAddress = GetPortLabel(sourcePortLabelFormat, portMapping.SourcePort);
        var destFormattedPortAddress = GetPortLabel(destinationPortLabelFormat, portMapping.DestinationPort);
        var value = $"{srcFormattedPortAddress}{Environment.NewLine} - {Environment.NewLine}{destFormattedPortAddress}";

        return new CellToAdd
        {
            Row = row,
            Column = 1,
            Value = value,
            Wrap = true
        };
    }


    private string GetPortLabel(string sourcePortLabelFormat, string[] portValues)
    {
        int index = 0;
        return Regex.Replace(sourcePortLabelFormat, @"\{[^}]*\}", match =>
        {
            if (index < portValues.Length)
            {
                return portValues[index++];
            }
            return match.Value; // If there are not enough replacements, keep the original value
        });
    }

    private IList<CellToAdd> GetPortCells(
        string[] portValues,
        int row,
        int columnOffset)
    {
        return portValues
            .Select((p, ix) => new CellToAdd { Row = row, Column = ix + columnOffset, Value = p })
            .ToList();
    }
}

public interface IRackCellFactory
{
    IList<CellToAdd> GetHeaderCells(RackSourceModel sourceModel);

    IList<CellToAdd> GetPortMappingCells(RackSourceModel sourceModel,
        int startRow,
        string sourcePortLabelFormat,
        string destinationPortLabelForma);
}
