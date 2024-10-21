using ClosedXML.Excel;
using Introl.Tools.Common.Enums;
using Introl.Tools.Common.Models;
using Introl.Tools.Timesheets.Models;
using Introl.Tools.Timesheets.Team.Models;
using OneOf;

namespace Introl.Tools.Timesheets.Team.Services;

public class EmployeeEmployeeEmployeeEmployeeTimesheetProcessor(
    ITeamSourceReader teamSourceReader,
    ITeamResultWriter teamResultWriter) : IEmployeeEmployeeTimesheetProcessor
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

        var inputSheetModel = teamSourceReader.Process(workbook);
        var outputWorkbookBytes = teamResultWriter.Process(inputSheetModel);

        return new ProcessedResult { Name = GetFileName(inputSheetModel), WorkbookBytes = outputWorkbookBytes };
    }

    private string GetFileName(TeamParsedSourceModel teamSourceModel)
    {
        var dateFormat = "yyyy.MM.dd";
        return
            $"Weekly Timesheet - Introl.io {teamSourceModel.StartDate.ToString(dateFormat)} - {teamSourceModel.EndDate.ToString(dateFormat)}.xlsx";

    }
}

public interface IEmployeeEmployeeTimesheetProcessor
{
    OneOf<ProcessedResult, ProcessingError> ProcessTimesheet(ProcessTimesheetRequest request);
}
