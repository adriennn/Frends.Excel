using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace Frends.Excel.ConvertToCSV;

public class Options
{
    /// <summary>
    /// If empty, all work sheets are read.
    /// </summary>
    [DefaultValue(@"")]
    public string ReadOnlyWorkSheetWithName { get; set; }

    /// <summary>
    /// Csv separator.
    /// </summary>
    [DefaultValue(@";")]
    [DisplayFormat(DataFormatString = "Text")]
    public string CsvSeparator { get; set; }

    /// <summary>
    /// If set to true, numbers will be used as column headers instead of letters (A = 1, B = 2...).
    /// </summary>
    [DefaultValue("false")]
    public bool UseNumbersAsColumnHeaders { get; set; }

    /// <summary>
    /// Choose if exception should be thrown when conversion fails.
    /// </summary>
    [DefaultValue("true")]
    public bool ThrowErrorOnFailure { get; set; }

    /// <summary>
    /// Date format selection.
    /// </summary>
    [DisplayName("Date Format")]
    [DisplayFormat(DataFormatString = "Text")]
    [DefaultValue(DateFormats.DEFAULT)]
    public DateFormats DateFormat { get; set; }

    /// <summary>
    /// If set to true, dates will exclude timestamps from dates.
    /// Default false
    /// </summary>
    [DefaultValue("false")]
    public bool ShortDatePattern { get; set; }
}
