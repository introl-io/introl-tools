using Swashbuckle.AspNetCore.Annotations;

namespace Introl.Timesheets.Api.Models;

public class CreateLabelsRequest
{
    [SwaggerSchema(Description = "File must be a .csv, or .xlsx file with the port mappings in the first sheet")]
    public required IFormFile File { get; init; }
    [SwaggerSchema(Description = "Format to use for source port in label. Put column letter or number in {} to be replaced with the column value. eg. {A}.R{2}.{E}")]
    public required string SourcePortLabelFormat { get; init; }
    [SwaggerSchema(Description = "Format to use for source port in label. Put column letter or number in {} to be replaced with the column value. eg. {A}.R{2}.{E}")]
    public required string DestinationPortLabelFormat { get; init; }
    public required bool HasHeadingRow { get; init; } = true;
}
