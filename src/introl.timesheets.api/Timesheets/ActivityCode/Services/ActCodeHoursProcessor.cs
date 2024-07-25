using Introl.Timesheets.Api.Timesheets.ActivityCode.Models;

namespace Introl.Timesheets.Api.Timesheets.ActivityCode.Services;

public class ActCodeHoursProcessor : IActCodeHoursProcessor
{
    public Dictionary<DateOnly, Dictionary<string, (double regHours, double otHours)>> Process(List<ActCodeHours> hours)
    {
        var result = new Dictionary<DateOnly, Dictionary<string, (double regHours, double otHours)>>();

        // var processedHours = 0;

        hours.Sort((a, b) => DateTime.Compare(a.StartTime, b.StartTime));
        
        foreach (var hr in hours)
        {
            var date = DateOnly.FromDateTime(hr.StartTime);

            if (result.TryGetValue(date, out var dayHours))
            {
                if (dayHours.TryGetValue(hr.ActivityCode, out var actCodeHrs))
                {
                    actCodeHrs.regHours += hr.Hours;
                    dayHours[hr.ActivityCode] = actCodeHrs;
                }
                else
                {
                    dayHours[hr.ActivityCode] = (hr.Hours, 0);
                }
            }
            else
            {
                result[date] = new Dictionary<string, (double regHours, double otHours)> { { hr.ActivityCode, (hr.Hours, 0) } };
            }
        }

        return result;
    }
}

public interface IActCodeHoursProcessor
{
    Dictionary<DateOnly, Dictionary<string, (double regHours, double otHours)>> Process(List<ActCodeHours> hours);
}
