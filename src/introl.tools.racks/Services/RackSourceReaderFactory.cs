namespace Introl.Tools.Racks.Services;

public class RackSourceReaderFactory(IEnumerable<IRackSourceReader> sourceReaders) : IRackSourceReaderFactory
{
    public IRackSourceReader? GetReader(string fileType)
    {
        return sourceReaders.FirstOrDefault(r => r.SupportedFileType == fileType);
    }
}

public interface IRackSourceReaderFactory
{
    IRackSourceReader? GetReader(string fileType);
}
