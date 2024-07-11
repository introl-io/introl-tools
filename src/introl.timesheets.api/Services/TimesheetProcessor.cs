using ClosedXML.Excel;
using Introl.Timesheets.Api.Enums;
using Introl.Timesheets.Api.Models;
using OneOf;

namespace Introl.Timesheets.Api.Services;

public class TimesheetProcessor(
    IWorksheetReader worksheetReader,
    IWorksheetWriter worksheetWriter) : ITimesheetProcessor
{
    public OneOf<ProcessedTimesheetResult, TimesheetProcessingFailureReasons> ProcessTimesheet(IFormFile inputFile)
    {
        var extension = Path.GetExtension(inputFile.FileName);
        if (extension != ".xlsx")
        {
            return TimesheetProcessingFailureReasons.UnsupportedFileType;
        }
        using var workbook = new XLWorkbook(inputFile.OpenReadStream());

        var inputSheetModel = worksheetReader.Process(workbook);
        var outputWorkbookBytes = worksheetWriter.Process(inputSheetModel);

        var dateFormat = "yyyy.MM.dd";
        var fileName =
            $"Weekly Timesheet - Introl.io {inputSheetModel.StartDate.ToString(dateFormat)} - {inputSheetModel.EndDate.ToString(dateFormat)}.xlsx";

        return new ProcessedTimesheetResult { Name = fileName, WorkbookBytes = outputWorkbookBytes };
    }
}

public interface ITimesheetProcessor
{
    OneOf<ProcessedTimesheetResult, TimesheetProcessingFailureReasons> ProcessTimesheet(IFormFile inputFile);
}
