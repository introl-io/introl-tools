using ClosedXML.Excel;
using Introl.Timesheets.Api.Constants;
using Introl.Timesheets.Api.Models;
using Introl.Timesheets.Api.Models.EmployeeTimesheets;

namespace Introl.Timesheets.Api.Services.EmployeeTimesheets;

public class EmployeeTimehsheetWriter(IOutputCellFactory outputCellFactory) : IEmployeeTimehsheetWriter
{
    public byte[] Process(EmployeeInputSheetModel employeeInputSheetModel)
    {
        using var workbook = new XLWorkbook();
        CreateSummarySheet(workbook, employeeInputSheetModel);

        workbook.AddWorksheet(employeeInputSheetModel.RawTimesheetsWorksheet);

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    private void CreateSummarySheet(XLWorkbook workbook, EmployeeInputSheetModel employeeInputSheetModel)
    {
        var employeeRow = 6;
        var worksheet = workbook.Worksheets.Add("Summary");
        var titleCells = outputCellFactory.GetTitleCells(worksheet, employeeInputSheetModel);

        var employeeCells = outputCellFactory.GetEmployeeCells(employeeInputSheetModel.Employees, ref employeeRow);

        var totalsCells = outputCellFactory.GetTotalsCells(employeeInputSheetModel, employeeRow + 4, employeeRow);

        WriteCells(worksheet, [.. titleCells, .. employeeCells, .. totalsCells]);

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
            worksheet.Cell(cell.Row, cell.Column).Style.Font.FontSize = cell.FontSize;
        }
    }
}

public interface IEmployeeTimehsheetWriter
{
    byte[] Process(EmployeeInputSheetModel employeeInputSheetModel);
}
