using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.Text;
using ExcelDataReader;
using Frends.Excel.ConvertToCSV.Definitions;

namespace Frends.Excel.ConvertToCSV;
public static class Excel
{
    /// <summary>
    /// Converts Excel file to CSV. [Documentation](https://github.com/FrendsPlatform/Frends.Excel/tree/main/Frends.Excel.ConvertToCSV)
    /// </summary>
    /// <param name="input">Input configuration</param>
    /// <param name="options">Input options</param>
    /// <param name="cancellationToken"></param>
    /// <returns>Result containing the converted CSV string.</returns>
    /// <exception cref="Exception"></exception>
    public static Result ConvertToCSV(
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
                    var csv = ConvertDataSetToCSV(result, options, cancellationToken);
                    return new Result(true, csv, null);
                }
            }
        }
        catch (Exception ex)
        {
            if (options.ThrowErrorOnFailure)
            {
                throw new Exception("Error while converting Excel file to CSV", ex);
            }

            return new Result(false, null, $"Error while converting Excel file to CSV: {ex}");
        }
    }

    private static string ConvertDataSetToCSV(DataSet result, Options options, CancellationToken cancellationToken)
    {
        var resultData = new StringBuilder();

        foreach (DataTable table in result.Tables)
        {
            // Read only wanted worksheets. If none is specified read all.
            if (options.ReadOnlyWorkSheetWithName.Contains(table.TableName) || options.ReadOnlyWorkSheetWithName.Length == 0)
            {
                for (var i = 0; i < table.Rows.Count; i++)
                {
                    for (var j = 0; j < table.Columns.Count; j++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        var item = table.Rows[i].ItemArray[j];
                        if (table.Rows[i].ItemArray[j].GetType().Name == "DateTime")
                        {
                            item = ConvertDateTimes((DateTime)item, options);
                        }
                        resultData.Append(item + options.CsvSeparator);
                    }
                    // Remove last CsvSeparator.
                    resultData.Length--;
                    resultData.Append(Environment.NewLine);
                }
            }
        }
        return resultData.ToString();
    }

    private static string ConvertDateTimes(DateTime date, Options options)
    {
        // Modify the date using date format var in options.

        if (options.ShortDatePattern)
        {
            switch (options.DateFormat)
            {
                case DateFormats.DDMMYYYY:
                    return date.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                case DateFormats.MMDDYYYY:
                    return date.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
                case DateFormats.YYYYMMDD:
                    return date.ToString("yyyy/MM/dd", CultureInfo.InvariantCulture);
                case DateFormats.DEFAULT:
                    return date.ToString(CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern);
                default:
                    return date.ToString(CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern);
            }
        }
        else
        {
            switch (options.DateFormat)
            {
                case DateFormats.DDMMYYYY:
                    return date.ToString("dd/MM/yyyy H:mm:ss", CultureInfo.InvariantCulture);
                case DateFormats.MMDDYYYY:
                    return date.ToString("MM/dd/yyyy h:mm:ss tt", CultureInfo.InvariantCulture);
                case DateFormats.YYYYMMDD:
                    return date.ToString("yyyy/MM/dd H:mm:ss", CultureInfo.InvariantCulture);
                case DateFormats.DEFAULT:
                    return date.ToString(CultureInfo.CurrentCulture.DateTimeFormat);
                default:
                    return date.ToString(CultureInfo.CurrentCulture.DateTimeFormat);
            }
        }
    }
}
