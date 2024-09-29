using Introl.Tools.Common.Utils;
using Introl.Tools.Racks.Models;

namespace Introl.Tools.Racks.Services;

public class RackSourceCsvReader : IRackSourceReader
{
    public string SupportedFileType => ".csv";
    public RackSourceModel Process(ProcessFileRequest request)
    {
        var csvText = new StreamReader(request.File.OpenReadStream()).ReadToEnd();
        var csvRows = csvText.Split(Environment.NewLine);

        return ParseRows(csvRows, request);
    }

    private RackSourceModel ParseRows(string[] csvRows,
        ProcessFileRequest request)
    {
        var sourceColumnsInts = GetZeroIndexedColumns(request.SourceColumns);
        var destColumnsInts = GetZeroIndexedColumns(request.DestinationColumns);
        List<string> srcPortHeadings = Enumerable.Repeat(string.Empty, sourceColumnsInts.Length).ToList();
        List<string> destPortHeadings = Enumerable.Repeat(string.Empty, destColumnsInts.Length).ToList();

        if (request.HasHeadingRow)
        {
            var headings = csvRows[0].Split(',');
            srcPortHeadings = sourceColumnsInts
                .Select(c => headings[c])
                .ToList();

            destPortHeadings = destColumnsInts
                .Select(c => headings[c])
                .ToList();
        }

        var startRow = request.HasHeadingRow ? 1 : 0;
        var portMappings = new List<RacksSourcePortMappingModel>();
        for (var i = startRow; i < csvRows.Length; i++)
        {
            var row = csvRows[i].Split(',');
            portMappings.Add(new RacksSourcePortMappingModel
            {
                SourcePort = GetPortModel(row, sourceColumnsInts),
                DestinationPort = GetPortModel(row, destColumnsInts),
            });
        }

        return new RackSourceModel
        {
            PortMappings = portMappings,
            SourcePortHeadings = srcPortHeadings,
            DestinationPortHeadings = destPortHeadings
        };
    }

    private int[] GetZeroIndexedColumns(string[] columns)
    {
        return columns.Select(c =>
        {
            if (int.TryParse(c, out var column))
            {
                return column - 1;
            }

            return ExcelUtils.ExcelColumnLetterToZeroBasedInt(c);
        }).ToArray();
    }



    private string[] GetPortModel(string[] row,
        int[] portColumns)
    {
        return portColumns.Select(column => row[column]).ToArray();
    }

}
