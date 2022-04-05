using System.ComponentModel;
using System.Text;
using ExcelDataReader;
using Frends.Excel.Parse.Definitions;

namespace Frends.Excel.Parse;

public static class Excel
{
    /// <summary>
    /// Converts Excel file to data set. [Documentation](https://tasks.frends.com/tasks#frends-tasks/Frends.Excel.Parse)
    /// </summary>
    /// <param name="input">Input configuration</param>
    /// <param name="options">Input options</param>
    /// <param name="cancellationToken"></param>
    /// <returns>Result containing the parsed Excel: object { bool Success, string ErrorMessage, DataSet DataSet }</returns>
    /// <exception cref="Exception"></exception>
    public static Result Parse(
        [PropertyTab] Input input,
        [PropertyTab] Options options,
        CancellationToken cancellationToken)
    {
        try
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            using (var stream = new FileStream(input.Path, FileMode.Open))
            {
                using (var excelReader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = excelReader.AsDataSet();
                    return new Result(true, result, null);
                }
            }
        }
        catch (Exception ex)
        {
            if (options.ThrowErrorOnFailure)
            {
                throw new Exception("Error while parsing Excel file", ex);
            }

            return new Result(false, null, $"Error while parsing Excel file: {ex}");
        }
    }
}
