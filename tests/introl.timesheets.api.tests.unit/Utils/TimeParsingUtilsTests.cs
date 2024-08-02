using Introl.Timesheets.Api.Utils;
using Xunit;

namespace Introl.Timesheets.Api.Tests.Unit.Utils;

public class TimeParsingUtilsTests
{
    [Theory]
    [InlineData("", 0)]
    [InlineData("1", 1)]
    [InlineData("2:00", 2)]
    [InlineData("3:19", 3)]
    [InlineData("3:20", 3.5)]
    [InlineData("3:39", 3.5)]
    [InlineData("3:40", 4)]
    [InlineData("3:59", 4)]
    public void ConvertToRoundedHours_ReturnsExpectedOutput(string input, double expectedOutput)
    {
        var result = TimeParsingUtils.ConvertToRoundedHours(input);
        
        Assert.Equal(expectedOutput, result);
    }
}
