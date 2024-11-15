using Introl.Tools.Common.Utils;
using Xunit;

namespace Introl.Tools.Api.Tests.Unit.Utils;

public class ExcelUtilsTests
{
    [Theory]
    [InlineData(1, "A")]
    [InlineData(2, "B")]
    [InlineData(26, "Z")]
    [InlineData(27, "AA")]
    [InlineData(53, "BA")]
    public void ToExcelColumn(int input, string expectedOutput)
    {
        var result = ExcelUtils.ToExcelColumn(input);

        Assert.Equal(expectedOutput, result);
    }
}
