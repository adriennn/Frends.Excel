namespace Frends.Excel.ConvertToJSON.Definitions;

internal class Row
{
    public int RowNumber { get; set; }
    public List<Cell>? Cells { get; set; }

    public Row() { }
}

