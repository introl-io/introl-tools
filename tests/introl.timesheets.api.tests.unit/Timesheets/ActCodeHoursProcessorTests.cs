using FluentAssertions;
using Introl.Timesheets.Api.Timesheets.ActivityCode.Models;
using Introl.Timesheets.Api.Timesheets.ActivityCode.Services;
using Xunit;

namespace Introl.Timesheets.Api.Tests.Unit.Timesheets;

public class ActCodeHoursProcessorTests
{
    private readonly IActCodeHoursProcessor _sut = new ActCodeHoursProcessor();

    [Theory, ClassData(typeof(ActCodeHoursProcessorTestData))]
    public void Process_WhenCalled_ReturnsExpectedResult(string _, List<ActCodeHours> input, Dictionary<string, Dictionary<DateOnly, (double regHours, double otHours)>> expected)
    {
        var result = _sut.Process(input);

       result.Should().BeEquivalentTo(expected);
    }
}
