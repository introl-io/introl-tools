using ClosedXML.Excel;

namespace Introl.Timesheets.Api.Services.ActivityCodeTimesheets;

public class ActivityCodeTimesheetReader : IActivityCodeTimesheetReader
{
    public void Process(XLWorkbook workbook)
    {
        var summaryWorksheet = workbook.Worksheets.Worksheet("Summary");
    }
}

public interface  IActivityCodeTimesheetReader
{
    void Process(XLWorkbook workbook);
}
