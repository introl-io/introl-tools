using ClosedXML.Excel;
using Introl.Timesheets.Api.Enums;
using Introl.Timesheets.Api.Models.EmployeeTimesheets;
using OneOf;

namespace Introl.Timesheets.Api.Services.ActivityCodeTimesheets;

public class ActivityCodeTimesheetProcessor : IActivityCodeTimesheetProcessor
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

        
        return new ProcessedTimesheetError
        {
            FailureReason = TimesheetProcessingFailureReasons.UnsupportedFileType,
            Message = "Unsupported file type: .xlsx. Please upload a .xlsx file."
        };
    }
}

public interface IActivityCodeTimesheetProcessor
{
    OneOf<ProcessedTimesheetResult, ProcessedTimesheetError> ProcessTimesheet(IFormFile inputFile);
}
