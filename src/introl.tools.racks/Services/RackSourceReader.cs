using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Introl.Tools.Racks.Models;

namespace Introl.Tools.Racks.Services;

public class RackSourceReader : IRackSourceReader
{
    public RacksSourceModel Process(XLWorkbook workbook)
    {
        var computeIbWorksheet = workbook.Worksheet("Compute IB");
        var keyPositions = GetKeyPositions(computeIbWorksheet);
        throw new Exception();
    }

    private RacksSourceModel ParseRows(IXLWorksheet worksheet, RacksSourceKeyPositions keyPositions)
    {
        throw new Exception();
    }
    private RacksSourceKeyPositions GetKeyPositions(IXLWorksheet worksheet)
    {
        var racks = GetSourceAndDestinationCells(worksheet, "Rack");
        var rackUnits = GetSourceAndDestinationCells(worksheet, "RU");
        var slots = GetSourceAndDestinationCells(worksheet, "Slot");
        var ports = GetSourceAndDestinationCells(worksheet, "Port");

        return new RacksSourceKeyPositions
        {
            StartRow = racks.Source.Address.RowNumber,
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

    private (IXLCell Source, IXLCell Destination) GetSourceAndDestinationCells(IXLWorksheet worksheet, string cellName)
    {
        var cells = worksheet.CellsUsed().Where(c => c.Value.ToString().ToUpper() == cellName).ToList();

        if (cells.Count != 2)
        {
            throw new ArgumentException($"Incorrect number of cells found for {cellName}, found {cells.Count} cells");
        }

        return (cells[0], cells[1]);
    }
}

public interface IRackSourceReader
{
    public RacksSourceModel Process(XLWorkbook workbook);
}
