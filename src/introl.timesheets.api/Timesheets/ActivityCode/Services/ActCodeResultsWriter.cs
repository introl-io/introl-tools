using Introl.Timesheets.Api.Timesheets.ActivityCode.Models;

namespace Introl.Timesheets.Api.Timesheets.ActivityCode.Services;

public class ActCodeResultsWriter : IActCodeResultsWriter
{
    public byte[] Process(ActCodeParsedSourceModel teamSourceModel)
    {
        throw new NotImplementedException();
    }
}

public interface IActCodeResultsWriter
{
    byte[] Process(ActCodeParsedSourceModel teamSourceModel);
}