using System.ComponentModel;
using System.Data;

namespace Frends.Excel.Parse.Definitions;

public class Result
{
    /// <summary>
    /// False if conversion fails.
    /// </summary>
    [DefaultValue("false")]
    public bool Success { get; internal set; }

    /// <summary>
    /// Exception message.
    /// </summary>
    [DefaultValue("")]
    public string ErrorMessage { get; internal set; }

    /// <summary>
    /// Parsed Excel data set.
    /// </summary>
    /// <returns>String</returns>
    public DataSet DataSet { get; internal set; }

    public Result(bool success, DataSet dataSet, string errorMessage)
    {
        Success = success;
        DataSet = dataSet;
        ErrorMessage = errorMessage;
    }
}
