using System.Text.RegularExpressions;

namespace Introl.Tools.Racks.Models;

public record ProcessFileRequest
{
    public required IFormFile File { get; init; }
    public required string SourcePortLabelFormat { get; init; }
    public required string DestinationPortLabelFormat { get; init; }
    public required bool HasHeadingRow { get; init; }

    public required int? LineCharacterLimit { get; init; }

    private string ColumnCaptureRegex => @"\{([^}]*)\}";

    private string[]? _sourceColumnsCache;
    public string[] SourceColumns => _sourceColumnsCache ??= Regex.Matches(SourcePortLabelFormat, ColumnCaptureRegex)
        .Select(match => match.Groups[1].Value)
        .ToArray();


    private string[]? _destinationColumnsCache;
    public string[] DestinationColumns => _destinationColumnsCache ??= Regex.Matches(DestinationPortLabelFormat, ColumnCaptureRegex)
        .Select(match => match.Groups[1].Value)
        .ToArray();

}
