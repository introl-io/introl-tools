using ClosedXML.Excel;
using Introl.Tools.Common.Extensions;
using Introl.Tools.Racks.Models;

namespace Introl.Tools.Racks.Services;

public class RackResultsWriter(IRackCellFactory cellFactory) : IRackResultsWriter
{
    public byte[] Process(RackSourceModel sourceModel,
        string sourcePortLabelFormat,
        string destinationPortLabelFormat,
        int? lineCharacterLimit)
    {
        var titleCells = cellFactory.GetHeaderCells(sourceModel);
        var mappingCells = cellFactory.GetPortMappingCells(
            sourceModel, 2,
            sourcePortLabelFormat,
            destinationPortLabelFormat,
            lineCharacterLimit);

        using var workbook = new XLWorkbook();

        var worksheet = workbook.Worksheets.Add("Labels");

        worksheet.WriteCells([.. titleCells, .. mappingCells]);

        worksheet.Columns().AdjustToContents();
        worksheet.Rows().Height = 55;
        worksheet.Row(1).Height = 20;

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }
}

public interface IRackResultsWriter
{
    byte[] Process(RackSourceModel sourceModel,
        string sourcePortLabelFormat,
        string destinationPortLabelFormat,
        int? lineCharacterLimit);
}
