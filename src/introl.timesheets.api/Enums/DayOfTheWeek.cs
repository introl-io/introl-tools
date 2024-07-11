using Introl.Timesheets.Api.Attributes;

namespace Introl.Timesheets.Api.Enums;

public enum DayOfTheWeek
{
    [StringValue("Mon")]
    Monday,
    [StringValue("Tue")]
    Tuesday,
    [StringValue("Wed")]
    Wednesday,
    [StringValue("Thu")]
    Thursday,
    [StringValue("Fri")]
    Friday,
    [StringValue("Sat")]
    Saturday,
    [StringValue("Sun")]
    Sunday
}
