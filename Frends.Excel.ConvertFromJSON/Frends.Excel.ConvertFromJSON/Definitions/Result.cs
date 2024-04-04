using System.ComponentModel;

namespace Frends.Excel.ConvertFromJSON.Definitions;

/// <summary>
/// Result of the JSON to Excel conversion.
/// </summary>
public class Result
{
    /// <summary>
    /// Indicates if the conversion was successful.
    /// </summary>
    /// <example>true</example>
    [DefaultValue("false")]
    public bool Success { get; private set; }

    /// <summary>
    /// Exception message.
    /// </summary>
    /// <example>An error occurred...</example>
    [DefaultValue("")]
    public string? ErrorMessage { get; private set; }

    /// <summary>
    /// Path to file produced by JSON-conversion to Excel.
    /// </summary>
    /// <example>"/tmp/ExcelFromJson.xlsx"</example>
    public string? ExcelFilePath { get; private set; }

    internal Result(bool success, string? filepath, string? errorMessage)
    {
        Success = success;
        ExcelFilePath = filepath;
        ErrorMessage = errorMessage;
    }
}
