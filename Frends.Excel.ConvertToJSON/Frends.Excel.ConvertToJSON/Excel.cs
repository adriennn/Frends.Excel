using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.Text;
using ExcelDataReader;
using Frends.Excel.ConvertToXML.Definitions;

namespace Frends.Excel.ConvertToXML;

public static class Excel
{
    /// <summary>
    /// Converts Excel file to XML. [Documentation](https://github.com/FrendsPlatform/Frends.Excel/tree/main/Frends.Excel.ConvertToXML)
    /// </summary>
    /// <param name="input">Input configuration</param>
    /// <param name="options">Input options</param>
    /// <param name="cancellationToken"></param>
    /// <returns>Result containing the converted XML string.</returns>
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
                throw new Exception("Error while converting Excel file to XML", ex);
            }

            return new Result(false, null, $"Error while converting Excel file to XML: {ex}");
        }
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

            if (options.ReadOnlyWorkSheetWithName.Contains(dt.TableName) ||
                options.ReadOnlyWorkSheetWithName.Length == 0)
            {
                json.Append("{");
                json.Append($"\"name\": \"{dt.TableName}\",");
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
                            json.Append(WriteRowToJson(dt, i, options).ToString());
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

                if (result.Tables.IndexOf(dt) != result.Tables.Count - 1)
                {
                    json.Append("},");
                }
                else
                {
                    json.Append("}");
                }
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

        var content = WriteColumnToJson(dt, i, options).ToString();
        if (content.Equals("[]"))
        {
            return "empty";
        }

        rowJson.Append(content);

        return rowJson;
    }

    private static object WriteColumnToJson(DataTable dt, int i, Options options)
    {
        var columnJson = new StringBuilder();
        columnJson.Append("[");
        for (var j = 0; j < dt.Columns.Count; j++)
        {
            var content = dt.Rows[i].ItemArray[j];
            if (string.IsNullOrWhiteSpace(content.ToString()) == false)
            {
                if (content.GetType().Name == "DateTime")
                {
                    content = ConvertDateTimes((DateTime)content, options);
                }

                content = content.ToString();
                columnJson.Append("{");
                if (options.UseNumbersAsColumnHeaders)
                {
                    columnJson.Append($"\"{j + 1}\":");
                }
                else
                {
                    columnJson.Append($"\"{ColumnIndexToColumnLetter(j + 1)}\":");
                }

                columnJson.Append($"\"{content}\"");
                if (j != dt.Columns.Count - 1)
                {
                    columnJson.Append("},");
                }
                else
                {
                    columnJson.Append("}");
                }

            }

        }

        columnJson.Append("]");

        return columnJson;
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
