using ClosedXML.Excel;
using Introl.Timesheets.Api.Constants;
using Introl.Timesheets.Api.Extensions;
using Introl.Timesheets.Api.Timesheets.ActivityCode.Models;

namespace Introl.Timesheets.Api.Timesheets.ActivityCode.Services;

public class ActCodeResultsWriter
    (IActCodeResultCellFactory resultCellFactory) : IActCodeResultsWriter
{
    public byte[] Process(ActCodeParsedSourceModel sourceModel)
    {
        using var workbook = new XLWorkbook();

        var worksheet = workbook.Worksheets.Add("Summary");
        
        var row = 1;
        var titleCells = resultCellFactory.GetTitleCells(worksheet, sourceModel, ref row);
        var employeeCells = resultCellFactory.CreateEmployeeCells(sourceModel, ref row);

        worksheet.WriteCells([..titleCells, ..employeeCells]);

        worksheet.Columns().AdjustToContents();
        worksheet.Rows().AdjustToContents();

        worksheet.SheetView.ZoomScale = DimensionConstants.ZoomLevel;
        if (worksheet.Column(1).Width < DimensionConstants.ImageWidthInCharacters)
        {
            worksheet.Column(1).Width = DimensionConstants.ImageWidthInCharacters;
        }

        if (worksheet.Row(1).Height < DimensionConstants.ImageHeightInPoints)
        {
            worksheet.Row(1).Height = DimensionConstants.ImageHeightInPoints;
        }
        
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
