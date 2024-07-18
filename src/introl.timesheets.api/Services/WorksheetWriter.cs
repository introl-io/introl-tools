using ClosedXML.Excel;
using Introl.Timesheets.Api.Constants;
using Introl.Timesheets.Api.Models;

namespace Introl.Timesheets.Api.Services;

public class WorksheetWriter(IWorksheetWriterHelper worksheetWriterHelper) : IWorksheetWriter
{
    public byte[] Process(InputSheetModel inputSheetModel)
    {
        using var workbook = new XLWorkbook();
        CreateSummarySheet(workbook, inputSheetModel);

        workbook.AddWorksheet(inputSheetModel.RawTimesheetsWorksheet);

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    private void CreateSummarySheet(XLWorkbook workbook, InputSheetModel inputSheetModel)
    {
        var worksheet = workbook.Worksheets.Add("Summary");
        var x = worksheetWriterHelper.GetTitleCells(worksheet, inputSheetModel);
        WriteCells(worksheet, x);
        var employeeRow = 6;
        worksheetWriterHelper.AddEmployeeRows(worksheet, inputSheetModel.Employees, ref employeeRow);
        worksheetWriterHelper.AddTotals(worksheet, inputSheetModel, employeeRow + 4, employeeRow);

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
    }

    private void WriteCells(IXLWorksheet worksheet, IEnumerable<CellToAdd> cells)
    {
        foreach (var cell in cells)
        {
            if (cell.Value.HasValue)
            {
                if (cell.ValueType == CellToAdd.CellValueType.Formula)
                {
                    worksheet.Cell(cell.Row, cell.Column).FormulaA1 = cell.Value.Value.ToString();
                }
                else
                {
                    worksheet.Cell(cell.Row, cell.Column).Value = cell.Value.Value;
                }
            }

            worksheet.Cell(cell.Row, cell.Column).Style.NumberFormat.Format = cell.NumberFormat;
            worksheet.Cell(cell.Row, cell.Column).Style.Font.Bold = cell.Bold;
            worksheet.Cell(cell.Row, cell.Column).Style.Fill.BackgroundColor = cell.Color;
        }

    }
}

public interface IWorksheetWriter
{
    byte[] Process(InputSheetModel inputSheetModel);
}
