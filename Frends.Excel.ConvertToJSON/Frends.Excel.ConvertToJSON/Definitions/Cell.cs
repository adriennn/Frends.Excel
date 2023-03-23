namespace Frends.Excel.ConvertToJSON.Definitions;

internal class Cell
{
    public dynamic? ColumnName { get; set; }
    public int ColumnIndex { get; set; }
    public string? ColumnValue { get; set; }

    public Cell() { }
}

