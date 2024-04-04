using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.Text;
using Newtonsoft.Json;
using ExcelDataReader;
using Frends.Excel.ConvertFromJSON.Definitions;
using DocumentFormat.OpenXml

namespace Frends.Excel.ConvertFromJSON;

/// <summary>
/// JSON to Excel converter task.
/// </summary>
public static class Excel
{
    /// <summary>
    /// Converts JSON to an Excel file. [Documentation](https://tasks.frends.com/tasks/frends-tasks/Frends.Excel.ConvertFromJSON)
    /// </summary>
    /// <param name="input">Input configuration</param>
    /// <param name="options">Input options</param>
    /// <param name="cancellationToken"></param>
    /// <returns>Result containing the converted JSON string.</returns>
    /// <exception cref="Exception"></exception>
    public static Result ConvertFromJSON(
        [PropertyTab] Input input,
        [PropertyTab] Options options,
        CancellationToken cancellationToken)
    {
        try
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            DataTable dt = (DataTable)JsonConvert.DeserializeObject(input.JSON, (typeof(DataTable)));
            // TODO, user should be able to provide output path
            var outputpath = ConvertDataSetToExcel(dt, options, cancellationToken);
            return new Result(true, outputpath, null);
        }
        catch (Exception ex)
        {
            if (options.ThrowErrorOnFailure)
                throw new InvalidOperationException("Error while converting JSON to Excel file", ex);

            return new Result(false, null, $"Error while converting JSON to Excel file: {ex}");
        }
    }

    private static dynamic ConvertDataSetToExcel(DataTable data, Options options, string fileName, CancellationToken cancellationToken)
    {
        var excelApp = OfficeOpenXML.GetInstance();
        using (var stream = excelApp.GetExcelStream(data, options.HAsHeaders)) // use true to hide datatable columns from excel
        {

            using (FileStream fs = new FileStream(options.WriteToPath, FileMode.Create))
            {
                stream.Seek(0, SeekOrigin.Begin);
                stream.CopyTo(fs);
                fs.Flush();
            }
        }

        return Options.WriteToPath;
    }

}

public sealed class OfficeOpenXML
{
    private static Lazy<OfficeOpenXML> _instance = new Lazy<OfficeOpenXML>(() => new OfficeOpenXML());
    private OfficeOpenXML(){}
    public static OfficeOpenXML GetInstance()
    {
        return _instance.Value;
    }

    public MemoryStream GetExcelStream(DataTable dt, bool firstRowAsHeader = true)
    {

        MemoryStream stream = new MemoryStream();
        using (var excel = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {

            WorkbookPart workbookPart = excel.AddWorkbookPart();
            Workbook workbook = new Workbook();
            Sheets sheets = new Sheets();

            var table = dt;
            // Create one worksheet
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            Worksheet worksheet = new Worksheet();
            SheetData data = new SheetData();
            List<Row> allRows = new List<Row>();

            // Set headers of the sheet
            Row headerRow = new Row() { RowIndex = 1 };
            for (int iColumn = 0; iColumn < table.Columns.Count; iColumn++)
            {
                var col = table.Columns[iColumn];
                // If first row of table is not the header then set columns of table as header of sheet
                if (!firstRowAsHeader)
                {
                    headerRow.Append(new Cell
                    {
                        DataType = CellValues.String,
                        CellValue = new CellValue(col.ColumnName)
                    });
                }
                else
                {
                    headerRow.Append(new Cell
                    {
                        DataType = CellValues.String,
                        CellValue = new CellValue(Convert.ToString(table.Rows[0][col]))
                    });
                }
            }

            allRows.Add(headerRow);

            // Append data rows
            for (int iRow = firstRowAsHeader ? 1 : 0; iRow < table.Rows.Count; iRow++)
            {
                var row = table.Rows[iRow];
                Row valueRow = new Row { RowIndex = (uint)(iRow + (firstRowAsHeader ? 1 : 2)) };

                for (int iColumn = 0; iColumn < table.Columns.Count; iColumn++)
                {
                    var col = table.Columns[iColumn];
                    valueRow.Append(new Cell
                    {
                        DataType = Format(col.DataType),
                        CellValue = new CellValue(Convert.ToString(row[col]))
                    });
                }
                allRows.Add(valueRow);
            }
            

            // Add rows to the worksheet
            data.Append(allRows);
            worksheet.Append(data);
            worksheetPart.Worksheet = worksheet;
            worksheetPart.Worksheet.Save();

            // Add worksheet to sheets
            Sheet sheet = new Sheet
            {
                Name = string.IsNullOrWhiteSpace(options.WriteToWorkSheetName) ? "Sheet1" : options.WriteToWorkSheetName,
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = 0
            }
            sheets.Append(sheet);
        }

