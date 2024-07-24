using ClosedXML.Excel;
using Introl.Timesheets.Api.Constants;
using Introl.Timesheets.Api.Extensions;
using Introl.Timesheets.Api.Models.ActivityCodeTimesheets;

namespace Introl.Timesheets.Api.Services.ActivityCodeTimesheets;

public class ActivityCodeTimesheetReader : IActivityCodeTimesheetReader
{
    public ActivityCodeTimesheetModel Process(XLWorkbook workbook)
    {
        var summaryWorksheet = workbook.Worksheets.Worksheet("Summary");
        
        var cols = GetActivityCodeSourceColumns(summaryWorksheet);
        throw new NotImplementedException();
    }
    
    private ActivityCodeSourceColumns GetActivityCodeSourceColumns(IXLWorksheet worksheet)
    {
        return new ActivityCodeSourceColumns
        {
            ActivityCode = worksheet.FindSingleCellByValue(ActivityCodeInputConstants.ActivityCode).Address.ColumnNumber,
            Date = worksheet.FindSingleCellByValue(ActivityCodeInputConstants.Date, true).Address.ColumnNumber,
            Name = worksheet.FindSingleCellByValue(ActivityCodeInputConstants.Name).Address.ColumnNumber,
            MemberCode = worksheet.FindSingleCellByValue(ActivityCodeInputConstants.MemberCode).Address.ColumnNumber,
            StartTime = worksheet.FindSingleCellByValue(ActivityCodeInputConstants.StartTime).Address.ColumnNumber,
            TrackedHours = worksheet.FindSingleCellByValue(ActivityCodeInputConstants.TrackedHours).Address.ColumnNumber,
            BillableRate = worksheet.FindSingleCellByValue(ActivityCodeInputConstants.BillableRate).Address.ColumnNumber,
        };
    }
}

public interface  IActivityCodeTimesheetReader
{
    ActivityCodeTimesheetModel Process(XLWorkbook workbook);
}
