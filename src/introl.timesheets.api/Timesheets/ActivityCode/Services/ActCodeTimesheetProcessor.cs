using ClosedXML.Excel;
using Introl.Timesheets.Api.Enums;
using Introl.Timesheets.Api.Models;
using OneOf;

namespace Introl.Timesheets.Api.Timesheets.ActivityCode.Services;

public class ActCodeTimesheetProcessor
    (IActCodeSourceReader timsheetReader) : IActCodeTimesheetProcessor
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
        var res = timsheetReader.Process(workbook);

        return new ProcessedTimesheetError
        {
            FailureReason = TimesheetProcessingFailureReasons.UnsupportedFileType,
            Message = "Unsupported file type: .xlsx. Please upload a .xlsx file."
        };
    }
}

public interface IActCodeTimesheetProcessor
{
    OneOf<ProcessedTimesheetResult, ProcessedTimesheetError> ProcessTimesheet(IFormFile inputFile);
}
