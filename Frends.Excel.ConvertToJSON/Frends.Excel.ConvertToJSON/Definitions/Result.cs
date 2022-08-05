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
    public bool Success { get; private set; }

    /// <summary>
    /// Exception message.
    /// </summary>
    /// <example>An error occurred...</example>
    [DefaultValue("")]
    public string? ErrorMessage { get; private set; }

    /// <summary>
    /// Excel-conversion to JSON.
    /// </summary>
    /// <example>{"Sheet1":[{"A1":"1","B1":"2"},{"A2":"3","B2":"4"}]}</example>
    public string? JSON { get; private set; }

    internal Result(bool success, string? json, string? errorMessage)
    {
        Success = success;
        JSON = json;
        ErrorMessage = errorMessage;
    }
}
