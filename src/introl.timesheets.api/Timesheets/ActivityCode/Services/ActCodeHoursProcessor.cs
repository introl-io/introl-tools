using Introl.Timesheets.Api.Timesheets.ActivityCode.Models;

namespace Introl.Timesheets.Api.Timesheets.ActivityCode.Services;

public class ActCodeHoursProcessor : IActCodeHoursProcessor
{
    public Dictionary<string, Dictionary<DateOnly, (double regHours, double otHours)>> Process(List<ActCodeHours> hours)
    {
        var result = new Dictionary<string, Dictionary<DateOnly, (double regHours, double otHours)>>();

        var processedHours = 0d;

        hours.Sort((a, b) => DateTime.Compare(a.StartTime, b.StartTime));

        foreach (var hr in hours)
        {
            var inOvertime = processedHours >= 40;

            var otHrs = inOvertime ? hr.Hours : 0d;
            var regHrs = inOvertime ? 0 : hr.Hours;

            if (!inOvertime && (processedHours + hr.Hours) > 40)
            {
                regHrs = 40 - processedHours;
                otHrs = hr.Hours - regHrs;
            }

            processedHours += hr.Hours;

            var date = DateOnly.FromDateTime(hr.StartTime);

            if (result.TryGetValue(hr.ActivityCode, out var dayHours))
            {
                if (dayHours.TryGetValue(date, out var actCodeHrs))
                {
                    actCodeHrs.regHours += regHrs;
                    actCodeHrs.otHours += otHrs;
                    dayHours[date] = actCodeHrs;
                }
                else
                {
                    dayHours[date] = (regHrs, otHrs);
                }
            }
            else
            {
                result[hr.ActivityCode] = new Dictionary<DateOnly, (double regHours, double otHours)> { { date, (regHrs, otHrs) } };
            }
        }

        return result;
    }
}

public interface IActCodeHoursProcessor
{
    Dictionary<string, Dictionary<DateOnly, (double regHours, double otHours)>> Process(List<ActCodeHours> hours);
}
