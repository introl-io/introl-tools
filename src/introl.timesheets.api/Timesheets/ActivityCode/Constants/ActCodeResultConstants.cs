namespace Introl.Timesheets.Api.Timesheets.ActivityCode.Constants;

public static class ActCodeResultConstants
{
    public static int NameColInt => 1;

    public static int EmployeeCodeInt => 2;
    public static int ActivityCodeColInt => 3;
    public static int HoursTypeColInt => 4;
    public static int DateStartColInt => 5;
    public static int TotalHoursColOffset => 1;
    public static int RateColOffset => 2;
    public static int TotalBillableColOffset => 3;
    public static int RegularHoursOffset => 0;
    public static int OtHoursOffset => 1;
    public static int PayrollHoursOffset => 2;
    public static int ActCodeTotalRows => 3;
    public static int TotalBlockTotalRows => 5; // Adding empty row between
}
