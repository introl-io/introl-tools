namespace Introl.Timesheets.Api.Models.ActivityCodeTimesheets;

public class ActivityCodeEmployee
{
    public required string Name { get; init; }    
    public required string MemberCode { get; init; } 
    public required List<ActivityCodeHours> ActivityCodeHours { get; init; }
}
