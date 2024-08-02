using ClosedXML.Excel;
using Introl.Timesheets.Api.Extensions;
using Introl.Timesheets.Api.Timesheets.ActivityCode.Models;

namespace Introl.Timesheets.Api.Timesheets.ActivityCode.Services;

public class ActCodeResultsWriter
    (IActCodeResultCellFactory resultCellFactory) : IActCodeResultsWriter
{
    public byte[] Process(ActCodeParsedSourceModel sourceModel)
    {
        var row = 1;
        var cells = resultCellFactory.CreateEmployeeCells(sourceModel, ref row);

        using var workbook = new XLWorkbook();

        var worksheet = workbook.Worksheets.Add("Summary");
        worksheet.WriteCells(cells);

        // workbook.AddWorksheet(sourceModel.InputWorksheet);

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }
}

public interface IActCodeResultsWriter
{
    byte[] Process(ActCodeParsedSourceModel sourceModel);
}
