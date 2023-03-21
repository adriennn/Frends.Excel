using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.Text;
using Newtonsoft.Json;
using ExcelDataReader;
using Frends.Excel.ConvertToJSON.Definitions;

namespace Frends.Excel.ConvertToJSON;

/// <summary>
/// Excel to JSON converter task.
/// </summary>
public static class Excel
{
    /// <summary>
    /// Converts Excel file to JSON. [Documentation](https://tasks.frends.com/tasks/frends-tasks/Frends.Excel.ConvertToJSON)
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

            using var stream = new FileStream(input.Path, FileMode.Open);
            using var excelReader = ExcelReaderFactory.CreateReader(stream);
            var result = excelReader.AsDataSet();
            var json = ConvertDataSetToJson(result, options, Path.GetFileName(input.Path), cancellationToken);
            return new Result(true, json, null);
        }
        catch (Exception ex)
        {
            if (options.ThrowErrorOnFailure)
                throw new InvalidOperationException("Error while converting Excel file to JSON", ex);

            return new Result(false, null, $"Error while converting Excel file to JSON: {ex}");
        }
    }

    private static dynamic ConvertDataSetToJson(DataSet result, Options options, string fileName, CancellationToken cancellationToken)
    {
        var sheets = new List<dynamic>();

        foreach (DataTable dt in result.Tables)
        {
            // If sheet name is specified AND current sheet isn't the one we want - skip
            if (!string.IsNullOrWhiteSpace(options.ReadOnlyWorkSheetWithName)
                && options.ReadOnlyWorkSheetWithName.Trim() != dt.TableName)
                continue;

            var sheet = new { name = dt.TableName, rows = new List<Row>() };

            for (var i = 0; i < dt.Rows.Count; i++)
            {
                var row = new Row();
                row.Cells = CollectColumnsFromRow(dt, i, options, cancellationToken);
                row.RowNumber = i + 1;
                if (row.Cells.Count > 0)
                    sheet.rows.Add(row);
            }
            sheets.Add(sheet);
        }

        object output = string.IsNullOrEmpty(options.ReadOnlyWorkSheetWithName)
            ? new { workbook = new { workbook_name = fileName, worksheets = sheets } }
            : new { workbook = new { workbook_name = fileName, worksheet = sheets[0] } };

        return JsonConvert.SerializeObject(output);
    }

    private static List<Cell> CollectColumnsFromRow(DataTable dt, int i, Options options, CancellationToken cancellationToken)
    {
        var columnValues = new List<Cell>();
        for (var j = 0; j < dt.Columns.Count; j++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var content = dt.Rows[i].ItemArray[j];

            if (content == null) continue;
            if (string.IsNullOrEmpty(content.ToString())) continue;

            if (content.GetType().Name == "DateTime")
                content = ConvertDateTimes((DateTime)content, options);

            columnValues.Add(new Cell 
            {
                ColumnName = options.UseNumbersAsColumnHeaders ? j + 1 : ColumnIndexToColumnLetter(j + 1),
                ColumnIndex = j + 1,
                ColumnValue = content.ToString()
            });
        }
        return columnValues;
    }

    private static string ConvertDateTimes(DateTime date, Options options)
    {
        if (options.ShortDatePattern)
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
