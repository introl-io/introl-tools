using ClosedXML.Excel;
using Introl.Timesheets.Api.constants;
using Introl.Timesheets.Api.models;

namespace Introl.Timesheets.Api.services;

public class WorksheetWriter(IWorksheetWriterHelper worksheetWriterHelper) : IWorksheetWriter
{
    public void Process(InputSheetModel inputSheetModel, XLWorkbook workbook)
    {
        CreateSummarySheet(workbook, inputSheetModel);
        
        workbook.AddWorksheet(inputSheetModel.RawTimesheetsWorksheet);
    }

    private void CreateSummarySheet(XLWorkbook workbook, InputSheetModel inputSheetModel)
    {
        var worksheet = workbook.Worksheets.Add("Summary");
        worksheetWriterHelper.AddTitleRows(worksheet, inputSheetModel);
        var employeeRow = 6;
        worksheetWriterHelper.AddEmployeeRows(worksheet, inputSheetModel.Employees, ref employeeRow);
        worksheetWriterHelper.AddTotals(worksheet, inputSheetModel, employeeRow + 4);
        
        worksheet.Columns().AdjustToContents();
        worksheet.Rows().AdjustToContents();

        worksheet.SheetView.ZoomScale = DimensionConstants.ZoomLevel;
        if(worksheet.Column(1).Width < DimensionConstants.ImageWidthInCharacters)
        {
            worksheet.Column(1).Width = DimensionConstants.ImageWidthInCharacters;
        }
        
        if(worksheet.Row(1).Height < DimensionConstants.ImageHeightInPoints)
        {
            worksheet.Row(1).Height = DimensionConstants.ImageHeightInPoints;
        }
    }
    
}

public interface IWorksheetWriter
{
    void Process(InputSheetModel inputSheetModel, XLWorkbook workbook);
}
