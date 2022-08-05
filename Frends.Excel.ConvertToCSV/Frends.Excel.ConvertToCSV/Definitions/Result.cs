using System.ComponentModel;
namespace Frends.Excel.ConvertToCSV.Definitions;

/// <summary>
/// Result.
/// </summary>
public class Result
{
    /// <summary>
    /// Conversion's status. False if conversion fails.
    /// </summary>
    /// <example>false</example>
    [DefaultValue("false")]
    public bool Success { get; private set; }

    /// <summary>
    /// Excel-conversion to CSV.
    /// </summary>
    /// <example>"Foo,Bar,Kanji 働,Summa\r\n1,2,3,6\r\nKuva,1,2,3\r\n,,,\r\nFoo,,,\r\n,Bar,,\r\n"</example>
    public string? CSV { get; private set; }

    /// <summary>
    /// Exception message.
    /// </summary>
    /// <example>Error while converting Excel file to CSV</example>
    public string? ErrorMessage { get; private set; }

    internal Result(bool success, string? csv, string? errorMessage)
    {
        Success = success;
        CSV = csv;
        ErrorMessage = errorMessage;
    }
}