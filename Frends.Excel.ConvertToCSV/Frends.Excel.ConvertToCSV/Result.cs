using System.ComponentModel;

namespace Frends.Excel.ConvertToCSV;

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
    public string CSV { get; private set; }

    public Result(bool success, string csv, string errorMessage)
    {
        Success = success;
        CSV = csv;
        ErrorMessage = errorMessage;
    }
}
