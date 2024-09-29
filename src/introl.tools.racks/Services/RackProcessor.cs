using ClosedXML.Excel;
using Introl.Tools.Common.Enums;
using Introl.Tools.Common.Models;
using Introl.Tools.Racks.Models;
using OneOf;

namespace Introl.Tools.Racks.Services;

public class RackProcessor(IRackSourceReader sourceReader, IRackResultsWriter resultsWriter) : IRackProcessor
{
    public OneOf<ProcessedResult, ProcessingError> Process(ProcessFileRequest request)
    {
        var supportedExtensions = new[] { ".xlsx", ".csv" };
        var extension = Path.GetExtension(request.File.FileName);
        if (!supportedExtensions.Contains(extension))
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
            request.DestinationPortLabelFormat);

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
