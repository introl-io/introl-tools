using ClosedXML.Excel;
using Introl.Tools.Common.Enums;
using Introl.Tools.Common.Models;
using OneOf;

namespace Introl.Tools.Racks.Services;

public class RackProcessor(IRackSourceReader sourceReader, IRackResultsWriter resultsWriter) : IRackProcessor
{
    public OneOf<ProcessedResult, ProcessingError> Process(IFormFile inputFile)
    {
        var extension = Path.GetExtension(inputFile.FileName);
        if (extension != ".xlsx")
        {
            return new ProcessingError
            {
                FailureReason = ProcessingFailureReasons.UnsupportedFileType,
                Message = $"Unsupported file type: {extension}. Please upload a .xlsx file."
            };
        }
        using var workbook = new XLWorkbook(inputFile.OpenReadStream());

        var inputSheetModel = sourceReader.Process(workbook);

        var resultFile = resultsWriter.Process(inputSheetModel);
        
        return new ProcessedResult
        {
            WorkbookBytes = resultFile,
            Name = "PortLabels.xlsx"
        };
    }
}

public interface IRackProcessor
{
    OneOf<ProcessedResult, ProcessingError> Process(IFormFile inputFile);
}
