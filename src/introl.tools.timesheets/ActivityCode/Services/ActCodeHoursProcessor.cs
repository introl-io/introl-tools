using Introl.Tools.Timesheets.ActivityCode.Models;

namespace Introl.Tools.Timesheets.ActivityCode.Services;

public class ActCodeHoursProcessor : IActCodeHoursProcessor
{
    public Dictionary<string, Dictionary<DateOnly, (double regHours, double otHours)>> Process(List<ActCodeHours> hours, bool calculateOvertime)
    {
        var result = new Dictionary<string, Dictionary<DateOnly, (double regHours, double otHours)>>();

        var processedHours = 0d;

        hours.Sort((a, b) => DateTime.Compare(a.StartTime, b.StartTime));

        foreach (var hr in hours)
        {
            var (regHrs, otHrs) = CalculateHours(hr.Hours, calculateOvertime, ref processedHours);
            
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
    
    private (double regHours, double otHours) CalculateHours(double hours, bool calculateOvertime, ref double processedHours)
    {
        if (!calculateOvertime)
        {
            return (hours, 0);
        }
        var inOvertime = processedHours >= 40;

        var otHrs = inOvertime ? hours : 0d;
        var regHrs = inOvertime ? 0 : hours;

        if (!inOvertime && (processedHours + hours) > 40)
        {
            regHrs = 40 - processedHours;
            otHrs = hours - regHrs;
        }
        processedHours += hours;
        return (regHrs, otHrs);
    }
}

public interface IActCodeHoursProcessor
{
    Dictionary<string, Dictionary<DateOnly, (double regHours, double otHours)>> Process(List<ActCodeHours> hours, bool calculateOvertime);
}
