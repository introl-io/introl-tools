using ClosedXML.Excel;
using Introl.Tools.Common.Enums;
using Introl.Tools.Common.Models;
using Introl.Tools.Racks.Models;
using OneOf;

namespace Introl.Tools.Racks.Services;

public class RackProcessor(IRackSourceReaderFactory sourceReaderFactory, IRackResultsWriter resultsWriter) : IRackProcessor
{
    public OneOf<ProcessedResult, ProcessingError> Process(ProcessFileRequest request)
    {
        var extension = Path.GetExtension(request.File.FileName);
        var sourceReader = sourceReaderFactory.GetReader(extension);
        if (sourceReader is null)
        {
            return new ProcessingError
            {
                FailureReason = ProcessingFailureReasons.UnsupportedFileType,
                Message = $"Unsupported file type: {extension}. Please upload a .xlsx file."
            };
        }

        var inputSheetModel = sourceReader.Process(request);

        var resultFile = resultsWriter.Process(
            inputSheetModel,
            request.SourcePortLabelFormat,
            request.DestinationPortLabelFormat,
            request.LineCharacterLimit);

        return new ProcessedResult
        {
            WorkbookBytes = resultFile,
            Name = "PortLabels.xlsx"
        };
    }
}

public interface IRackProcessor
{
    OneOf<ProcessedResult, ProcessingError> Process(ProcessFileRequest request);
}
