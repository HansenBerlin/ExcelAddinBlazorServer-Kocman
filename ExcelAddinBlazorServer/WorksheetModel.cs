namespace ExcelAddinBlazorServer;

public class WorksheetModel
{
    public string? SheetName { get; set; }
    public TableModel[]? Tables { get; set; }
}