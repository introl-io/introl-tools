using ClosedXML.Excel;
using Introl.Tools.Racks.Models;

namespace Introl.Tools.Racks.Services;

public class RackSourceReader : IRackSourceReader
{
    public RackSourceModel Process(XLWorkbook workbook)
    {
        var computeIbWorksheet = workbook.Worksheet("Compute IB");
        var keyPositions = GetKeyPositions(computeIbWorksheet);

        return ParseRows(computeIbWorksheet, keyPositions);
    }

    private RackSourceModel ParseRows(IXLWorksheet worksheet, RacksSourceKeyPositions keyPositions)
    {
        var portMappings = new List<RacksSourcePortMappingModel>();
        for (var i = keyPositions.StartRow; i <= keyPositions.EndRow; i++)
        {
            var testCell = worksheet.Cell(i, 1).GetString();
            if (string.IsNullOrWhiteSpace(testCell))
            {
                continue;
            }
            portMappings.Add(new RacksSourcePortMappingModel
            {
                SourcePort = GetPortModel(worksheet, i, keyPositions.SourcePort),
                DestinationPort = GetPortModel(worksheet, i, keyPositions.DestinationPort),
            });
        }

        return new RackSourceModel
        {
            PortMappings = portMappings
        };
    }

    private RacksSourcePortMappingModel.PortModel GetPortModel(IXLWorksheet worksheet, int row,
        RacksSourceKeyPositions.PortKeyPositions portPositions)
    {
        return new RacksSourcePortMappingModel.PortModel
        {
            Rack = worksheet.Cell(row, portPositions.RackColumn).GetString().ToUpper(),
            RackUnit = $"R{worksheet.Cell(row, portPositions.RackUnitColumn).GetString()}".ToUpper(),
            Port = worksheet.Cell(row, portPositions.PortColumn).GetString().ToUpper(),
            Switch = worksheet.Cell(row, portPositions.SlotColumn).GetString().ToUpper(),
        };
    }

    private RacksSourceKeyPositions GetKeyPositions(IXLWorksheet worksheet)
    {
        var racks = GetSourceAndDestinationTitleCells(worksheet, "Rack");
        var rackUnits = GetSourceAndDestinationTitleCells(worksheet, "RU");
        var slots = GetSourceAndDestinationTitleCells(worksheet, "Slot");
        var ports = GetSourceAndDestinationTitleCells(worksheet, "Port");

        return new RacksSourceKeyPositions
        {
            StartRow = racks.Source.Address.RowNumber + 1,
            EndRow = worksheet.RowsUsed().Last().RowNumber(),
            SourcePort = new RacksSourceKeyPositions.PortKeyPositions
            {
                RackColumn = racks.Source.Address.ColumnNumber,
                RackUnitColumn = rackUnits.Source.Address.ColumnNumber,
                SlotColumn = slots.Source.Address.ColumnNumber,
                PortColumn = ports.Source.Address.ColumnNumber,
            },
            DestinationPort = new RacksSourceKeyPositions.PortKeyPositions
            {
                RackColumn = racks.Destination.Address.ColumnNumber,
                RackUnitColumn = rackUnits.Destination.Address.ColumnNumber,
                SlotColumn = slots.Destination.Address.ColumnNumber,
                PortColumn = ports.Destination.Address.ColumnNumber,
            }
        };
    }

    private (IXLCell Source, IXLCell Destination) GetSourceAndDestinationTitleCells(IXLWorksheet worksheet, string cellName)
    {
        var cells = worksheet.Row(1).CellsUsed().Where(c => c.GetString().ToUpper() == cellName.ToUpper()).ToList();

        if (cells.Count != 2)
        {
            throw new ArgumentException($"Incorrect number of cells found for {cellName}, found {cells.Count} cells");
        }

        return (cells[0], cells[1]);
    }
}

public interface IRackSourceReader
{
    public RackSourceModel Process(XLWorkbook workbook);
}
