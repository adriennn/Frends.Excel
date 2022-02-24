using System.ComponentModel;

namespace Frends.Excel.ConvertToXML;

public class Result
{
    private readonly string _csv;

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
    /// Excel-conversion to CSV.
    /// </summary>
    /// <returns>String</returns>
    public string XML { get; private set; }

    public Result(bool success, string xml, string errorMessage)
    {
        Success = success;
        XML = xml;
        ErrorMessage = errorMessage;
    }
}
