using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace Frends.Excel.ConvertToJSON.Definitions;

/// <summary>
/// Input paramameters for the Excel to JSON conversion.
/// </summary>
public class Input
{
    /// <summary>
    /// Path to the Excel file.
    /// </summary>
    /// <example>C:\tmp\ExcelFile.xlsx</example>
    [DefaultValue(@"C:\tmp\ExcelFile.xlsx")]
    [DisplayFormat(DataFormatString = "Text")]
    public string Path { get; set; } = "";
}
