using System.ComponentModel;

namespace Frends.Excel.ConvertToXML.Definitions;

public class Result
{
    /// <summary>
    /// False if conversion fails.
    /// </summary>
    [DefaultValue("false")]
    public bool Success { get; set; }

    /// <summary>
    /// Exception message.
    /// </summary>
    [DefaultValue("")]
    public string ErrorMessage { get; private set; }

    /// <summary>
    /// Excel-conversion to JSON.
    /// </summary>
    /// <returns>String</returns>
    public string JSON { get; private set; }

    public Result(bool success, string json, string errorMessage)
    {
        Success = success;
        JSON = json;
        ErrorMessage = errorMessage;
    }
}
