using System.ComponentModel;

namespace Frends.Excel.ConvertToXML.Definitions;

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
    public bool Success { get; set; }

    /// <summary>
    /// Excel-conversion to CSV.
    /// </summary>
    /// <example>workbook_name, worksheet_name, row_header, column_header</example>
    public string? XML { get; private set; }

    /// <summary>
    /// Exception message.
    /// </summary>
    /// <example>Error while converting Excel file to XML</example>
    public string? ErrorMessage { get; private set; }


    internal Result(bool success, string? xml, string? errorMessage)
    {
        Success = success;
        XML = xml;
        ErrorMessage = errorMessage;
    }
}