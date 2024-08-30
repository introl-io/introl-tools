namespace Introl.Tools.Timesheets.Utils;

public static class TimeParsingUtils
{
    public static double ConvertToRoundedHours(string inputHours)
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
            >= 0 and <= 19 => hours,
            >= 20 and <= 39 => hours + 0.5,
            _ => hours + 1
        };
    }
}
