using Introl.Timesheets.Api.Timesheets.ActivityCode.Models;

namespace Introl.Timesheets.Api.Timesheets.ActivityCode.Services;

public class ActCodeHoursProcessor : IActCodeHoursProcessor
{
    public Dictionary<DateOnly, Dictionary<string, (double regHours, double otHours)>> Process(List<ActCodeHours> hours)
    {
        var result = new Dictionary<DateOnly, Dictionary<string, (double regHours, double otHours)>>();

        var processedHours = 0d;

        hours.Sort((a, b) => DateTime.Compare(a.StartTime, b.StartTime));
        
        foreach (var hr in hours)
        {
            var inOvertime = processedHours >= 40;
            
            var otHrs = inOvertime ? hr.Hours : 0d;
            var regHrs = inOvertime ? 0 : hr.Hours;
            
            if(!inOvertime && (processedHours + hr.Hours) > 40)
            {
                regHrs = 40 - processedHours;
                otHrs = hr.Hours - regHrs;
            }
            
            processedHours += hr.Hours;
            
            var date = DateOnly.FromDateTime(hr.StartTime);

            if (result.TryGetValue(date, out var dayHours))
            {
                if (dayHours.TryGetValue(hr.ActivityCode, out var actCodeHrs))
                {
                    actCodeHrs.regHours += regHrs;
                    actCodeHrs.regHours += otHrs;
                    dayHours[hr.ActivityCode] = actCodeHrs;
                }
                else
                {
                    dayHours[hr.ActivityCode] = (regHrs, otHrs);
                }
            }
            else
            {
                result[date] = new Dictionary<string, (double regHours, double otHours)> { { hr.ActivityCode, (regHrs, otHrs) } };
            }
        }

        return result;
    }
}

public interface IActCodeHoursProcessor
{
    Dictionary<DateOnly, Dictionary<string, (double regHours, double otHours)>> Process(List<ActCodeHours> hours);
}
