using ClosedXML.Excel;
using Introl.Timesheets.Api.Enums;
using Introl.Timesheets.Api.Models;
using Introl.Timesheets.Api.Timesheets.ActivityCode.Models;
using OneOf;

namespace Introl.Timesheets.Api.Timesheets.ActivityCode.Services;

public class ActCodeTimesheetProcessor(
    IActCodeSourceReader timesheetReader,
    IActCodeResultsWriter resultsWriter) : IActCodeTimesheetProcessor
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
        var sourceModel = timesheetReader.Process(workbook);
        var resultBytes = resultsWriter.Process(sourceModel);
        return new ProcessedTimesheetResult
        {
            WorkbookBytes = resultBytes,
            Name = GetFileName(sourceModel)
        };
    }

    private string GetFileName(ActCodeParsedSourceModel sourceModel)
    {
        var dateFormat = "yyyy.MM.dd";
        return
            $"Weekly Timesheet - Introl.io {sourceModel.StartDate.ToString(dateFormat)} - {sourceModel.EndDate.ToString(dateFormat)}.xlsx";

    }
}

public interface IActCodeTimesheetProcessor
{
    OneOf<ProcessedTimesheetResult, ProcessedTimesheetError> ProcessTimesheet(IFormFile inputFile);
}
