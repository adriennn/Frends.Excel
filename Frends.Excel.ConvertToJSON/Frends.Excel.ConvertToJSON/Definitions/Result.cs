using System.ComponentModel;

namespace Frends.Excel.ConvertToJSON.Definitions;

/// <summary>
/// Result of the Excel to JSON conversion.
/// </summary>
public class Result
{
    /// <summary>
    /// Indicates if the conversion was successful.
    /// </summary>
    /// <example>true</example>
    [DefaultValue("false")]
    public bool Success { get; set; }

    /// <summary>
    /// Exception message.
    /// </summary>
    [DefaultValue("")]
    public string? ErrorMessage { get; private set; }

    /// <summary>
    /// Excel-conversion to JSON.
    /// </summary>
    public string? JSON { get; private set; }

    internal Result(bool success, string? json, string? errorMessage)
    {
        Success = success;
        JSON = json;
        ErrorMessage = errorMessage;
    }
}
