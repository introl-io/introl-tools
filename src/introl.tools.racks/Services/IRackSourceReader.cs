using Introl.Tools.Racks.Models;

namespace Introl.Tools.Racks.Services;

public interface IRackSourceReader
{
    public string SupportedFileType { get; }
    public RackSourceModel Process(ProcessFileRequest request);
}
