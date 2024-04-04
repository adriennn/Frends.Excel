using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace Frends.Excel.ConvertFromJSON.Definitions;

/// <summary>
/// Input paramameters for the JSON to Excel conversion.
/// </summary>
public class Input
{
    /// <summary>
    /// JToken representation of a JSON object.
    /// </summary>
    /// <example>...</example>
    [DefaultValue(@"...")]
    [DisplayFormat(DataFormatString = "Text")]
    public string JSON { get; set; } = "";
}
