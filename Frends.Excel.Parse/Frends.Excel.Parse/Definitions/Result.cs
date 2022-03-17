using System.ComponentModel;
using System.Data;

namespace Frends.Excel.Parse.Definitions;

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
    /// Parsed Excel data set.
    /// </summary>
    /// <returns>String</returns>
    public DataSet DataSet { get; set; }

    public Result(bool success, DataSet dataSet, string errorMessage)
    {
        Success = success;
        DataSet = dataSet;
        ErrorMessage = errorMessage;
    }
}
