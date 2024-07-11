using ClosedXML.Excel;
using Introl.Timesheets.Api.Enums;
using Introl.Timesheets.Api.Models;
using OneOf;

namespace Introl.Timesheets.Api.Services;

public class TimesheetProcessor(
    IWorksheetReader worksheetReader,
    IWorksheetWriter worksheetWriter) : ITimesheetProcessor
{
    public OneOf<ProcessedTimesheetResult, ProcessedTimesheetError> ProcessTimesheet(IFormFile inputFile)
    {
        var extension = Path.GetExtension(inputFile.FileName);
        if (extension != ".xlsx")
        {
            return new ProcessedTimesheetError
            {
                FailureReason = TimesheetProcessingFailureReasons.UnsupportedFileType,
                Message = $"Unsupported file type: {extension}. Please upload a .xlsx file."
            };
        }
        using var workbook = new XLWorkbook(inputFile.OpenReadStream());

        var inputSheetModel = worksheetReader.Process(workbook);
        var outputWorkbookBytes = worksheetWriter.Process(inputSheetModel);
        
        return new ProcessedTimesheetResult { Name = GetFileName(inputSheetModel), WorkbookBytes = outputWorkbookBytes };
    }

    private string GetFileName(InputSheetModel inputSheetModel)
    {
        var dateFormat = "yyyy.MM.dd";
        return
            $"Weekly Timesheet - Introl.io {inputSheetModel.StartDate.ToString(dateFormat)} - {inputSheetModel.EndDate.ToString(dateFormat)}.xlsx";

    }
}

public interface ITimesheetProcessor
{
    OneOf<ProcessedTimesheetResult, ProcessedTimesheetError> ProcessTimesheet(IFormFile inputFile);
}
