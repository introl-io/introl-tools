using Introl.Timesheets.Api.Extensions;
using Xunit;

namespace Introl.Timesheets.Api.Tests.Unit.Extensions;

public class IntExtensionsTests
{
    [Theory]
    [InlineData(1, "A")]
    [InlineData(2, "B")]
    [InlineData(26, "Z")]
    [InlineData(27, "AA")]
    [InlineData(53, "BA")]
    public void ToExcelColumn(int input, string expectedOutput)
    {
        var result = input.ToExcelColumn();
        
        Assert.Equal(expectedOutput, result);
    }
}
