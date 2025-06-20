namespace OfficeOpenXml.Style;

internal interface IStyle
{
	ulong Id { get; }

	ExcelStyle ExcelStyle { get; }

	void SetNewStyleID(string value);
}
