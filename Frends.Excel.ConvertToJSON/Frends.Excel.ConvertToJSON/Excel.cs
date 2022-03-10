using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.Text;
using ExcelDataReader;
using Frends.Excel.ConvertToJSON.Definitions;

namespace Frends.Excel.ConvertToJSON;

public static class Excel
{
    /// <summary>
    /// Converts Excel file to JSON. [Documentation](https://github.com/FrendsPlatform/Frends.Excel/tree/main/Frends.Excel.ConvertToJSON)
    /// </summary>
    /// <param name="input">Input configuration</param>
    /// <param name="options">Input options</param>
    /// <param name="cancellationToken"></param>
    /// <returns>Result containing the converted JSON string.</returns>
    /// <exception cref="Exception"></exception>
    public static Result ConvertToJSON(
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
                    var json = ConvertDataSetToJson(result, options, Path.GetFileName(input.Path));
                    return new Result(true, json, null);
                }
            }
        }
        catch (Exception ex)
        {
            if (options.ThrowErrorOnFailure)
            {
                throw new Exception("Error while converting Excel file to JSON", ex);
            }

            return new Result(false, null, $"Error while converting Excel file to JSON: {ex}");
        }
    }

    private static string SanitizeJSONValue(string input)
    {
        return input.Replace("\"", "\\\"");
    }

    private static string ConvertDataSetToJson(DataSet result, Options options, string fileName)
    {
        var json = new StringBuilder();
        json.Append("{");
        json.Append($"\"workbook\": ");
        json.Append("{");
        json.Append($"\"workbook_name\": \"{fileName}\",");
        if (options.ReadOnlyWorkSheetWithName.Length == 0)
        {
            json.Append("\"worksheets\": ");
            json.Append("[");
        }
        else
        {
            json.Append("\"worksheet\" : ");
        }

        foreach (DataTable dt in result.Tables)
        {
            // If sheet name is specified AND current sheet isn't the one we want - skip
            if (!string.IsNullOrWhiteSpace(options.ReadOnlyWorkSheetWithName)
                && options.ReadOnlyWorkSheetWithName.Trim() != dt.TableName)
            {
                continue;
            }

            json.Append("{");
            json.Append($"\"name\": \"{SanitizeJSONValue(dt.TableName)}\",");
            json.Append("\"rows\": ");

            // Building json from datatable.
            if (dt.Rows.Count > 0)
            {
                json.Append("[");
                for (var i = 0; i < dt.Rows.Count; i++)
                {
                    var content = WriteRowToJson(dt, i, options).ToString();
                    if (!content.ToString().Equals("empty"))
                    {
                        json.Append(content);
                        if (i < dt.Rows.Count - 1)
                        {
                            json.Append("},");
                        }
                        else if (i == dt.Rows.Count - 1)
                        {
                            json.Append("}");
                        }
                    }
                }
                json.Append("]");
            }

            json.Append("}");

            // Append comma when this is either the last sheet
            // An exception is when we serialize only one sheet - then we never need the comma (thus check for that param)
            if (string.IsNullOrWhiteSpace(options.ReadOnlyWorkSheetWithName)
                && result.Tables.IndexOf(dt) != result.Tables.Count - 1)
            {
                json.Append(",");
            }
        }

        if (options.ReadOnlyWorkSheetWithName.Length == 0)
        {
            json.Append("]");
        }

        json.Append("}");
        json.Append("}");

        return json.ToString();
    }

    private static object WriteRowToJson(DataTable dt, int i, Options options)
    {
        var rowJson = new StringBuilder();
        rowJson.Append("{");
        rowJson.Append($"\"{i + 1}\":");

        var content = WriteRowColumnsToJson(dt, i, options);
        if (content.Equals("[]"))
        {
            return "empty";
        }

        rowJson.Append(content);

        return rowJson;
    }

    private static string WriteRowColumnsToJson(DataTable dt, int i, Options options)
    {
        var columnValues = new List<string>();
        for (var j = 0; j < dt.Columns.Count; j++)
        {
            var content = dt.Rows[i].ItemArray[j];
            if (string.IsNullOrWhiteSpace(content.ToString())) continue;

            if (content.GetType().Name == "DateTime")
            {
                content = ConvertDateTimes((DateTime)content, options);
            }

            content = SanitizeJSONValue(content.ToString());

            var columnHeader = options.UseNumbersAsColumnHeaders
                ? $"\"{j + 1}\""
                : $"\"{ColumnIndexToColumnLetter(j + 1)}\"";

            var columnValue = $"{{{columnHeader}:\"{content}\"}}";
            columnValues.Add(columnValue);
        }

        return $"[{string.Join(',', columnValues)}]";
    }

    private static string ConvertDateTimes(DateTime date, Options options)
    {
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

    private static string ColumnIndexToColumnLetter(int colIndex)
    {
        var div = colIndex;
        var colLetter = string.Empty;
        int mod;
        while (div > 0)
        {
            mod = (div - 1) % 26;
            colLetter = (char)(65 + mod) + colLetter;
            div = ((div - mod) / 26);
        }

        return colLetter;
    }
}
