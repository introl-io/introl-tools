using ClosedXML.Excel;
using Introl.Timesheets.Api.Enums;
using Introl.Timesheets.Api.Models;
using Introl.Timesheets.Api.Timesheets.Team.Models;
using OneOf;

namespace Introl.Timesheets.Api.Timesheets.Team.Services;

public class EmployeeEmployeeEmployeeEmployeeTimesheetProcessor(
    ITeamSourceReader teamSourceReader,
    ITeamResultWriter teamResultWriter) : IEmployeeEmployeeTimesheetProcessor
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

        var inputSheetModel = teamSourceReader.Process(workbook);
        var outputWorkbookBytes = teamResultWriter.Process(inputSheetModel);

        return new ProcessedTimesheetResult { Name = GetFileName(inputSheetModel), WorkbookBytes = outputWorkbookBytes };
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
    OneOf<ProcessedTimesheetResult, ProcessedTimesheetError> ProcessTimesheet(IFormFile inputFile);
}
