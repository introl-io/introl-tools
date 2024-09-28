using System.Text.RegularExpressions;
using ClosedXML.Excel;
using Introl.Tools.Racks.Models;

namespace Introl.Tools.Racks.Services;

public class RackSourceReader : IRackSourceReader
{
    public RackSourceModel Process(
        XLWorkbook workbook,
        string sourcePortLabelFormat,
        string destinationPortLabelFormat,
        bool hasHeadingRow)
    {
        var worksheet = workbook.Worksheets.First();
        var keyPositions = GetKeyPositions(worksheet, sourcePortLabelFormat, destinationPortLabelFormat, hasHeadingRow);

        return ParseRows(worksheet, keyPositions, hasHeadingRow);
    }

    private RackSourceModel ParseRows(
        IXLWorksheet worksheet,
        RacksSourceKeyPositions keyPositions,
        bool hasHeadingRow)
    {
        List<string> srcPortHeadings = Enumerable.Repeat(string.Empty, keyPositions.SourceColumns.Length).ToList();
        List<string> destPortHeadings = Enumerable.Repeat(string.Empty, keyPositions.DestinationColumns.Length).ToList();

        if (hasHeadingRow)
        {
            srcPortHeadings = keyPositions.SourceColumns
                .Select(c => worksheet.Cell(1, c).GetString())
                .ToList();

            destPortHeadings = keyPositions.DestinationColumns
                .Select(c => worksheet.Cell(1, c).GetString())
                .ToList();
        }

        var portMappings = new List<RacksSourcePortMappingModel>();
        for (var i = keyPositions.StartRow; i <= keyPositions.EndRow; i++)
        {
            var testCell = worksheet.Cell(i, keyPositions.SourceColumns[0]).GetString();
            if (string.IsNullOrWhiteSpace(testCell))
            {
                continue;
            }

            portMappings.Add(new RacksSourcePortMappingModel
            {
                SourcePort = GetPortModel(worksheet, i, keyPositions.SourceColumns),
                DestinationPort = GetPortModel(worksheet, i, keyPositions.DestinationColumns),
            });
        }

        return new RackSourceModel
        {
            PortMappings = portMappings,
            SourcePortHeadings = srcPortHeadings,
            DestinationPortHeadings = destPortHeadings
        };
    }

    private string[] GetPortModel(IXLWorksheet worksheet, int row,
        string[] portColumns)
    {
        return portColumns.Select(column => worksheet.Cell(row, column).GetString().ToUpper()).ToArray();
    }

    private RacksSourceKeyPositions GetKeyPositions(
        IXLWorksheet worksheet,
        string sourcePortLabelFormat,
        string destinationPortLabelFormat,
        bool hasHeadingRow)
    {
        return new RacksSourceKeyPositions
        {
            StartRow = hasHeadingRow ? 2 : 1,
            EndRow = worksheet.RowsUsed().Last().RowNumber(),
            SourceColumns = ExtractColumns(sourcePortLabelFormat),
            DestinationColumns = ExtractColumns(destinationPortLabelFormat)
        };
    }

    private string[] ExtractColumns(string labelFormat)
    {
        return Regex.Matches(labelFormat, @"\{([^}]*)\}")
            .Cast<Match>()
            .Select(match => match.Groups[1].Value)
            .ToArray();
    }
}

public interface IRackSourceReader
{
    public RackSourceModel Process(
        XLWorkbook workbook,
        string sourcePortLabelFormat,
        string destinationPortLabelFormat,
        bool hasHeadingRow);
}
