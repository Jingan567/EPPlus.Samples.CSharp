namespace OfficeOpenXml;

public class ExcelTableAddress
{
	public string Name { get; set; }

	public string ColumnSpan { get; set; }

	public bool IsAll { get; set; }

	public bool IsHeader { get; set; }

	public bool IsData { get; set; }

	public bool IsTotals { get; set; }

	public bool IsThisRow { get; set; }
}
