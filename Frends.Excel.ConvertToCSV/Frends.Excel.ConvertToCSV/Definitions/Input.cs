using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace Frends.Excel.ConvertToCSV.Definitions;

/// <summary>
/// Input parameters.
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
