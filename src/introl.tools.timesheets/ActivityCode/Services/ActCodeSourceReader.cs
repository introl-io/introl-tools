using ClosedXML.Excel;
using Introl.Tools.Common.Extensions;
using Introl.Tools.Timesheets.ActivityCode.Constants;
using Introl.Tools.Timesheets.ActivityCode.Models;
using Introl.Tools.Timesheets.Utils;

namespace Introl.Tools.Timesheets.ActivityCode.Services;

public class ActCodeSourceReader : IActCodeSourceReader
{
    public ActCodeParsedSourceModel Process(XLWorkbook workbook)
    {
        var summaryWorksheet = workbook.Worksheets.Worksheet("Summary");

        var (startDate, endDate) = GetStartAndEndDate(summaryWorksheet);
        var projectCode = summaryWorksheet.FindSingleCellByValue(ActCodeSourceConstants.ProjectCode).CellBelow().GetString();
        var keyPositions = GetActivityCodeKeyPositions(summaryWorksheet);
        var employees = GetEmployees(summaryWorksheet, keyPositions);
        var activityCodes = summaryWorksheet
            .Column(keyPositions.ActivityCodeCol)
            .CellsUsed()
            .Where(c => c.Address.RowNumber < keyPositions.TotalHoursRow)
            .Select(c => c.Value.ToString())
            .Where(c => c.ToUpper() != ActCodeSourceConstants.ActivityCode.ToUpper())
            .Distinct();
        return new ActCodeParsedSourceModel
        {
            StartDate = startDate,
            EndDate = endDate,
            InputWorksheet = summaryWorksheet,
            Employees = employees,
            ActivityCodes = activityCodes,
            ProjectCode = projectCode
        };
    }

    private IDictionary<string, ActCodeEmployee> GetEmployees(IXLWorksheet worksheet,
        ActCodeSourceKeyPositions keyPositions)
    {
        var startRow = keyPositions.TitleRow + 1;
        var endRow = keyPositions.TotalHoursRow - 1;
        var result = new Dictionary<string, ActCodeEmployee>();

        for (var i = startRow; i < endRow; i++)
        {
            var memberCode = worksheet.Cell(i, keyPositions.MemberCodeCol);
            var activityCode = worksheet.Cell(i, keyPositions.ActivityCodeCol);
            if (string.IsNullOrWhiteSpace(memberCode.GetString())
                || string.IsNullOrWhiteSpace(activityCode.GetString()))
            {
                continue;
            }

            var totalHours = TimeParsingUtils.ConvertToRoundedHours(worksheet.Cell(i, keyPositions.TrackedHoursCol).GetString());
            var date = worksheet.Cell(i, keyPositions.DateCol).GetDateTime();
            var startTime = worksheet.Cell(i, keyPositions.StartTimeCol).GetTimeSpan();
            var startDateTime = new DateTime(date.Year, date.Month, date.Day, startTime.Hours, startTime.Minutes, 0);

            var rate = worksheet.Cell(i, keyPositions.BillableRateCol).GetDouble();
            var hours = new ActCodeHours
            {
                StartTime = startDateTime,
                ActivityCode = worksheet.Cell(i, keyPositions.ActivityCodeCol).GetString(),
                Hours = totalHours
            };
            if (result.ContainsKey(memberCode.GetString()))
            {
                result[memberCode.GetString()].ActivityCodeHours.Add(hours);
            }
            else
            {
                result[memberCode.GetString()] = new ActCodeEmployee
                {
                    Name = worksheet.Cell(i, keyPositions.NameCol).GetString(),
                    Rate = rate,
                    MemberCode = memberCode.GetString(),
                    ActivityCodeHours = new List<ActCodeHours> { hours }
                };
            }
        }

        return result;
    }

    private ActCodeSourceKeyPositions GetActivityCodeKeyPositions(IXLWorksheet worksheet)
    {
        return new ActCodeSourceKeyPositions
        {
            TitleRow = worksheet.FindSingleCellByValue(ActCodeSourceConstants.ActivityCode, true).Address.RowNumber,
            ActivityCodeCol = worksheet.FindSingleCellByValue(ActCodeSourceConstants.ActivityCode, true).Address.ColumnNumber,
            DateCol = worksheet.FindSingleCellByValue(ActCodeSourceConstants.Date, true).Address.ColumnNumber,
            NameCol = worksheet.FindSingleCellByValue(ActCodeSourceConstants.Name).Address.ColumnNumber,
            MemberCodeCol = worksheet.FindSingleCellByValue(ActCodeSourceConstants.MemberCode).Address.ColumnNumber,
            StartTimeCol = worksheet.FindSingleCellByValue(ActCodeSourceConstants.StartTime).Address.ColumnNumber,
            TrackedHoursCol =
                worksheet.FindSingleCellByValue(ActCodeSourceConstants.TrackedHours).Address.ColumnNumber,
            BillableRateCol =
                worksheet.FindSingleCellByValue(ActCodeSourceConstants.BillableRate).Address.ColumnNumber,
            TotalHoursRow = worksheet.FindSingleCellByValue(ActCodeSourceConstants.TotalHours).Address.RowNumber,
        };
    }

    private (DateOnly startDate, DateOnly endDate) GetStartAndEndDate(IXLWorksheet worksheet)
    {
        var dateRangeCell = worksheet.FindSingleCellByValue(ActCodeSourceConstants.DateRange);
        var dateString = dateRangeCell.CellRight().GetString();
        var splitDates = dateString.Split(" - ");

        var startDate = DateOnly.Parse(splitDates[0]);
        var endDate = DateOnly.Parse(splitDates[1]);

        return (startDate, endDate);
    }
}

public interface IActCodeSourceReader
{
    ActCodeParsedSourceModel Process(XLWorkbook workbook);
}
