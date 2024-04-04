using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace Frends.Excel.ConvertFromJSON.Definitions;

/// <summary>
/// Options paramameters for the JSON to Excel conversion.
/// </summary>
public class Options
{
    /// <summary>
    /// If empty, write to Sheet1 as default name.
    /// </summary>
    /// <example>Sheet1</example>
    [DefaultValue(@"")]
    public string WriteToWorkSheetName { get; set; } = "";


    /// <summary>
    /// If set to true, the first row of the excel file will contain the JSON attributes as column names.
    /// </summary>
    /// <example>true</example>
    [DefaultValue("false")]
    public bool HAsHeaders { get; set; }

    /// <summary>
    /// If empty, write to file.xlsx file name in path ./tmp/.
    /// </summary>
    /// <example>./tmp/file.xlsx</example>
    [DefaultValue(@"")]
    public string WriteToPath { get; set; } = "";

    /// <summary>
    /// Choose if exception should be thrown when conversion fails.
    /// </summary>
    /// <example>true</example>
    [DefaultValue("true")]
    public bool ThrowErrorOnFailure { get; set; }

}
