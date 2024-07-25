using ClosedXML.Excel;
using Introl.Timesheets.Api.Enums;
using Introl.Timesheets.Api.Models.EmployeeTimesheets;
using OneOf;

namespace Introl.Timesheets.Api.Services.EmployeeTimesheets;

public class EmployeeEmployeeTimesheetProcessor(
    IEmployeeTimehsheetReader employeeTimehsheetReader,
    IEmployeeTimehsheetWriter employeeTimehsheetWriter) : IEmployeeTimesheetProcessor
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

        var inputSheetModel = employeeTimehsheetReader.Process(workbook);
        var outputWorkbookBytes = employeeTimehsheetWriter.Process(inputSheetModel);

        return new ProcessedTimesheetResult { Name = GetFileName(inputSheetModel), WorkbookBytes = outputWorkbookBytes };
    }

    private string GetFileName(EmployeeInputSheetModel employeeInputSheetModel)
    {
        var dateFormat = "yyyy.MM.dd";
        return
            $"Weekly Timesheet - Introl.io {employeeInputSheetModel.StartDate.ToString(dateFormat)} - {employeeInputSheetModel.EndDate.ToString(dateFormat)}.xlsx";

    }
}

public interface IEmployeeTimesheetProcessor
{
    OneOf<ProcessedTimesheetResult, ProcessedTimesheetError> ProcessTimesheet(IFormFile inputFile);
}
