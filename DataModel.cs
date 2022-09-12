namespace SpreadsheetReader;

public sealed class DataModel
{
    public DateOnly OrderDate { get; set; }
    public string Region { get; set; }
    public string Rep { get; set; }
    public string Item { get; set; }
    public int Units { get; set; }
    public decimal UnitCost { get; set; }
    public decimal Total { get; set; }
}
