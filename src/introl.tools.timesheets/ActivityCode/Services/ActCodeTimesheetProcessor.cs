using ClosedXML.Excel;
using Introl.Tools.Common.Enums;
using Introl.Tools.Common.Models;
using Introl.Tools.Timesheets.ActivityCode.Models;
using Introl.Tools.Timesheets.Models;
using OneOf;

namespace Introl.Tools.Timesheets.ActivityCode.Services;

public class ActCodeTimesheetProcessor(
    IActCodeSourceReader timesheetReader,
    IActCodeResultsWriter resultsWriter) : IActCodeTimesheetProcessor
{
    public OneOf<ProcessedResult, ProcessingError> ProcessTimesheet(ProcessTimesheetRequest request)
    {
        var extension = Path.GetExtension(request.File.FileName);
        if (extension != ".xlsx")
        {
            return new ProcessingError
            {
                FailureReason = ProcessingFailureReasons.UnsupportedFileType,
                Message = $"Unsupported file type: {extension}. Please upload a .xlsx file."
            };
        }

        using var workbook = new XLWorkbook(request.File.OpenReadStream());
        var sourceModel = timesheetReader.Process(workbook);
        var resultBytes = resultsWriter.Process(sourceModel);
        return new ProcessedResult
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
    OneOf<ProcessedResult, ProcessingError> ProcessTimesheet(ProcessTimesheetRequest request);
}
