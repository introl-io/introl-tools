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

        var (startDate, endDate) = GetStartAndEndDate(summaryWorksheet);
        var keyPositions = GetActivityCodeKeyPositions(summaryWorksheet);
        var employees = GetEmployees(summaryWorksheet, keyPositions);
        
        return new ActivityCodeTimesheetModel
        {
            StartDate = startDate,
            EndDate = endDate,
            InputWorksheet = summaryWorksheet,
            Employees = employees
        };
    }

    private IDictionary<string, ActivityCodeEmployee> GetEmployees(IXLWorksheet worksheet,
        ActivityCodeSourceKeyPositions keyPositions)
    {
        var startRow = keyPositions.TitleRow + 1;
        var endRow = keyPositions.TotalHoursRow - 1;
        var result = new Dictionary<string, ActivityCodeEmployee>();

        for (var i = startRow; i < endRow; i++)
        {
            var memberCode = worksheet.Cell(i, keyPositions.MemberCodeCol);
            if (string.IsNullOrWhiteSpace(memberCode.GetString()))
            {
                continue;
            }

            var totalHours = ConvertToRoundedHours(worksheet.Cell(i, keyPositions.TrackedHoursCol).GetString());
            var date = worksheet.Cell(i, keyPositions.DateCol).GetDateTime();
            var startTime = worksheet.Cell(i, keyPositions.StartTimeCol).GetDateTime();

            var startDateTime = new DateTime(date.Year, date.Month, date.Day, startTime.Hour, startTime.Minute, 0);

            var hours = new ActivityCodeHours
            {
                StartTime = startDateTime,
                ActivityCode = worksheet.Cell(i, keyPositions.ActivityCodeCol).GetString(),
                Hours = totalHours
            };
            if(result.ContainsKey(memberCode.GetString()))
            {
                result[memberCode.GetString()].ActivityCodeHours.Add(hours);
            }
            else
            {
                result[memberCode.GetString()] = new ActivityCodeEmployee
                {
                    Name = worksheet.Cell(i, keyPositions.NameCol).GetString(),
                    MemberCode = memberCode.GetString(),
                    ActivityCodeHours = new List<ActivityCodeHours> { hours }
                };
            }
        }

        return result;
    }

    private ActivityCodeSourceKeyPositions GetActivityCodeKeyPositions(IXLWorksheet worksheet)
    {
        return new ActivityCodeSourceKeyPositions
        {
            TitleRow = worksheet.FindSingleCellByValue(ActivityCodeInputConstants.ActivityCode).Address.RowNumber,
            ActivityCodeCol = worksheet.FindSingleCellByValue(ActivityCodeInputConstants.ActivityCode).Address.ColumnNumber,
            DateCol = worksheet.FindSingleCellByValue(ActivityCodeInputConstants.Date, true).Address.ColumnNumber,
            NameCol = worksheet.FindSingleCellByValue(ActivityCodeInputConstants.Name).Address.ColumnNumber,
            MemberCodeCol = worksheet.FindSingleCellByValue(ActivityCodeInputConstants.MemberCode).Address.ColumnNumber,
            StartTimeCol = worksheet.FindSingleCellByValue(ActivityCodeInputConstants.StartTime).Address.ColumnNumber,
            TrackedHoursCol =
                worksheet.FindSingleCellByValue(ActivityCodeInputConstants.TrackedHours).Address.ColumnNumber,
            BillableRateCol =
                worksheet.FindSingleCellByValue(ActivityCodeInputConstants.BillableRate).Address.ColumnNumber,
            TotalHoursRow = worksheet.FindSingleCellByValue(ActivityCodeInputConstants.TotalHours).Address.RowNumber,
        };
    }
    
    private double ConvertToRoundedHours(string inputHours)
    {
        if (!inputHours.Contains(":"))
        {
            if (double.TryParse(inputHours, out var parsedHours))
            {
                return parsedHours;
            }

            return 0;
        }

        var splitHours = inputHours.Split(':');
        var hours = double.Parse(splitHours[0]);
        var minutes = double.Parse(splitHours[1]);

        return minutes switch
        {
            >= 0 and <= 14 => hours,
            >= 15 and <= 44 => hours + 0.5,
            _ => hours + 1
        };
    }
    
    private (DateOnly startDate, DateOnly endDate) GetStartAndEndDate(IXLWorksheet worksheet)
    {
        var dateRangeCell = worksheet.FindSingleCellByValue(ActivityCodeInputConstants.DateRange);
        var dateString = dateRangeCell.CellRight().GetString();
        var splitDates = dateString.Split(" - ");

        var startDate = DateOnly.Parse(splitDates[0]);
        var endDate = DateOnly.Parse(splitDates[1]);

        return (startDate, endDate);
    }
}

public interface IActivityCodeTimesheetReader
{
    ActivityCodeTimesheetModel Process(XLWorkbook workbook);
}
