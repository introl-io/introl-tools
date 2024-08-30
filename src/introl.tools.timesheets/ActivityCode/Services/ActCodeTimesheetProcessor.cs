using ClosedXML.Excel;
using Introl.Tools.Timesheets.ActivityCode.Models;
using Introl.Tools.Timesheets.Enums;
using Introl.Tools.Timesheets.Models;
using OneOf;

namespace Introl.Tools.Timesheets.ActivityCode.Services;

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
            $"{sourceModel.ProjectCode} Timesheet - Introl.io {sourceModel.StartDate.ToString(dateFormat)} - {sourceModel.EndDate.ToString(dateFormat)}.xlsx";

    }
}

public interface IActCodeTimesheetProcessor
{
    OneOf<ProcessedTimesheetResult, ProcessedTimesheetError> ProcessTimesheet(IFormFile inputFile);
}
