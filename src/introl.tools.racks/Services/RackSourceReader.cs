using System.Text.RegularExpressions;
using ClosedXML.Excel;
using Introl.Tools.Racks.Models;

namespace Introl.Tools.Racks.Services;

public class RackSourceReader : IRackSourceReader
{
    public RackSourceModel Process(ProcessFileRequest request)
    {
        using var workbook = new XLWorkbook(request.File.OpenReadStream());
        var worksheet = workbook.Worksheets.First();

        return ParseRows(worksheet, request);
    }

    private RackSourceModel ParseRows(
        IXLWorksheet worksheet,
        ProcessFileRequest request)
    {
        List<string> srcPortHeadings = Enumerable.Repeat(string.Empty, request.SourceColumns.Length).ToList();
        List<string> destPortHeadings = Enumerable.Repeat(string.Empty, request.DestinationColumns.Length).ToList();

        if (request.HasHeadingRow)
        {
            srcPortHeadings = request.SourceColumns
                .Select(c => worksheet.Cell(1, c).GetString())
                .ToList();

            destPortHeadings = request.DestinationColumns
                .Select(c => worksheet.Cell(1, c).GetString())
                .ToList();
        }

        var startRow = request.HasHeadingRow ? 2 : 1;
        var endRow = worksheet.RowsUsed().Last().RowNumber();
        var portMappings = new List<RacksSourcePortMappingModel>();
        for (var i = startRow; i <= endRow; i++)
        {
            var testCell = worksheet.Cell(i, request.SourceColumns[0]).GetString();
            if (string.IsNullOrWhiteSpace(testCell))
            {
                continue;
            }

            portMappings.Add(new RacksSourcePortMappingModel
            {
                SourcePort = GetPortModel(worksheet, i, request.SourceColumns),
                DestinationPort = GetPortModel(worksheet, i, request.DestinationColumns),
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

}

public interface IRackSourceReader
{
    public RackSourceModel Process(ProcessFileRequest request);
}
