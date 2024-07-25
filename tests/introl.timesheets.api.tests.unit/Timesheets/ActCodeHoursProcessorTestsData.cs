using System.Collections;
using Introl.Timesheets.Api.Timesheets.ActivityCode.Models;

namespace Introl.Timesheets.Api.Tests.Unit.Timesheets;

public class ActCodeHoursProcessorTestData : IEnumerable<object[]>
{
    private readonly List<object[]> _data = new()
    {

        new object[] { FortyHrsInput, FortyHrsResult }

    };
    
    public IEnumerator<object[]> GetEnumerator()

    { return _data.GetEnumerator(); }

 

    IEnumerator IEnumerable.GetEnumerator()

    { return GetEnumerator(); }
    
    private static DateOnly StartDate => new(2021, 1, 1);

    private static List<ActCodeHours> FortyHrsInput =>
    [
        new ActCodeHours
        {
            StartTime = new DateTime(StartDate, new TimeOnly(9, 0)), 
            Hours = 5, 
            ActivityCode = "A"
        },
        new ActCodeHours
        {
            StartTime = new DateTime(StartDate.AddDays(4), new TimeOnly(8, 30)), 
            Hours = 10, 
            ActivityCode = "A"
        },
        new ActCodeHours
        {
            StartTime = new DateTime(StartDate, new TimeOnly(14, 0)), 
            Hours = 2, 
            ActivityCode = "B"
        },
        new ActCodeHours
        {
            StartTime = new DateTime(StartDate.AddDays(5), new TimeOnly(13, 0)), 
            Hours = 6, 
            ActivityCode = "B"
        },
        new ActCodeHours
        {
            StartTime = new DateTime(StartDate.AddDays(3), new TimeOnly(10, 0)), 
            Hours = 10, 
            ActivityCode = "B"
        },
        new ActCodeHours
        {
            StartTime = new DateTime(StartDate, new TimeOnly(16, 0)), 
            Hours = 7, 
            ActivityCode = "A"
        }
    ];

    private static Dictionary<DateOnly, Dictionary<string, (double regHours, double otHours)>> FortyHrsResult =>
        new()
        {
            {
                StartDate,
                new Dictionary<string, (double regHours, double otHours)> { { "A", (12, 0) }, { "B", (2, 0) } }
            },
            {
                StartDate.AddDays(3),
                new Dictionary<string, (double regHours, double otHours)> { { "B", (10, 0) } }
            },
            {
                StartDate.AddDays(4),
                new Dictionary<string, (double regHours, double otHours)> { { "A", (10, 0) } }
            },
            {
                StartDate.AddDays(5),
                new Dictionary<string, (double regHours, double otHours)> { { "B", (6, 0) } }
            }
        };
}
