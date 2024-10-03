using Swashbuckle.AspNetCore.Annotations;

namespace Introl.Timesheets.Api.Models;

public record CreateLabelsRequest
{
    [SwaggerSchema(Description = "File must be a .csv, or .xlsx file with the port mappings in the first sheet")]
    public required IFormFile File { get; init; }
    [SwaggerSchema(Description = "Format to use for source port in label. Put column letter or number in {} to be replaced with the column value. eg. {A}.R{2}.{E}")]
    public required string SourcePortLabelFormat { get; init; }
    [SwaggerSchema(Description = "Format to use for source port in label. Put column letter or number in {} to be replaced with the column value. eg. {A}.R{2}.{E}")]
    public required string DestinationPortLabelFormat { get; init; }
    [SwaggerSchema(Description = "Is the first row of the file a heading row")]
    public required bool HasHeadingRow { get; init; } = true;
    [SwaggerSchema(Description = "Optional: If set will limit the number of characters in a label line")]
    public int? LineCharacterLimit { get; init; } = null;
}
