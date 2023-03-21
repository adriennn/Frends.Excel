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
    /// <example>"{"workbook":{"workbook_name":"ExcelTestInput2.xls","worksheet":{"name":"Sheet1","rows":[{"RowNumber":1,"Cells":[{"ColumnName":"A","ColumnIndex":1,"ColumnValue":"Foo"},{"ColumnName":"B","ColumnIndex":2,"ColumnValue":"Bar"},{"ColumnName":"C","ColumnIndex":3,"ColumnValue":"Kanji働"},{"ColumnName":"D","ColumnIndex":4,"ColumnValue":"Summa"}]}]}}}"</example>
    public string? JSON { get; private set; }

    internal Result(bool success, string? json, string? errorMessage)
    {
        Success = success;
        JSON = json;
        ErrorMessage = errorMessage;
    }
}
