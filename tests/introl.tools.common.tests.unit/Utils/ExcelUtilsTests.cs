using Introl.Tools.Common.Utils;
using Xunit;

namespace Introl.Tools.Common.Tests.Unit.Utils;

public class ExcelUtilsTests
{
    [Theory]
    [InlineData("A", 0)]
    [InlineData("Z", 25)]
    [InlineData("AA", 26)]
    // [InlineData("BA", 26)]
    public void ExcelColumnLetterToZeroBasedInt_WhenCalled_ReturnsExpected(string column, int expected)
    {
        // Act
        var result = ExcelUtils.ExcelColumnLetterToZeroBasedInt(column);

        // Assert
        Assert.Equal(expected, result);
    }
}
