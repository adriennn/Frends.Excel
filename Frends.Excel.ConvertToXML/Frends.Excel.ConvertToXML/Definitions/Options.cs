using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace Frends.Excel.ConvertToXML.Definitions;

/// <summary>
/// Options.
/// </summary>
public class Options
{
    /// <summary>
    /// If empty, all work sheets are read.
    /// </summary>
    /// <example>Sheet2</example>
    [DisplayFormat(DataFormatString = "Text")]
    public string? ReadOnlyWorkSheetWithName { get; set; }

    /// <summary>
    /// If set to true, numbers will be used as column headers instead of letters (A = 1, B = 2...).
    /// </summary>
    /// <example>false</example>
    [DefaultValue("false")]
    public bool UseNumbersAsColumnHeaders { get; set; }

    /// <summary>
    /// Choose if exception should be thrown when conversion fails.
    /// </summary>
    /// <example>true</example>
    [DefaultValue("true")]
    public bool ThrowErrorOnFailure { get; set; }

    /// <summary>
    /// Date format selection.
    /// </summary>
    /// <example>Default</example>
    [DisplayName("Date Format")]
    [DisplayFormat(DataFormatString = "Text")]
    [DefaultValue(DateFormats.DEFAULT)]
    public DateFormats DateFormat { get; set; }

    /// <summary>
    /// If set to true, dates will exclude timestamps from dates.
    /// Default false
    /// </summary>
    /// <example>false</example>
    [DefaultValue("false")]
    public bool ShortDatePattern { get; set; }

    internal bool ShouldReadSheet(string sheetName)
    {
        if (string.IsNullOrWhiteSpace(ReadOnlyWorkSheetWithName))
            return true;
        if (ReadOnlyWorkSheetWithName.Equals(sheetName))
            return true;

        return false;
    }
}
