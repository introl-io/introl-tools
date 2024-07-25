using ClosedXML.Excel;
using Introl.Timesheets.Api.Constants;
using Introl.Timesheets.Api.Models;
using Introl.Timesheets.Api.Timesheets.Team.Models;

namespace Introl.Timesheets.Api.Timesheets.Team.Services;

public class TeamResultWriter(ITeamResultCellFactory teamResultCellFactory) : ITeamResultWriter
{
    public byte[] Process(TeamParsedSourceModel teamSourceModel)
    {
        using var workbook = new XLWorkbook();
        CreateSummarySheet(workbook, teamSourceModel);

        workbook.AddWorksheet(teamSourceModel.RawTimesheetsWorksheet);

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    private void CreateSummarySheet(XLWorkbook workbook, TeamParsedSourceModel teamSourceModel)
    {
        var employeeRow = 6;
        var worksheet = workbook.Worksheets.Add("Summary");
        var titleCells = teamResultCellFactory.GetTitleCells(worksheet, teamSourceModel);

        var employeeCells = teamResultCellFactory.GetEmployeeCells(teamSourceModel.Employees, ref employeeRow);

        var totalsCells = teamResultCellFactory.GetTotalsCells(teamSourceModel, employeeRow + 4, employeeRow);

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

public interface ITeamResultWriter
{
    byte[] Process(TeamParsedSourceModel teamSourceModel);
}
