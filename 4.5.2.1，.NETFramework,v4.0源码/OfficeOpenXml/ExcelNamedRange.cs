namespace OfficeOpenXml;

public sealed class ExcelNamedRange : ExcelRangeBase
{
	private ExcelWorksheet _sheet;

	public string Name { get; internal set; }

	public int LocalSheetId
	{
		get
		{
			if (_sheet == null)
			{
				return -1;
			}
			return _sheet.PositionID - _workbook._package._worksheetAdd;
		}
	}

	internal ExcelWorksheet LocalSheet => _sheet;

	internal int Index { get; set; }

	public bool IsNameHidden { get; set; }

	public string NameComment { get; set; }

	internal object NameValue { get; set; }

	internal string NameFormula { get; set; }

	public ExcelNamedRange(string name, ExcelWorksheet nameSheet, ExcelWorksheet sheet, string address, int index)
		: base(sheet, address)
	{
		Name = name;
		_sheet = nameSheet;
		Index = index;
	}

	internal ExcelNamedRange(string name, ExcelWorkbook wb, ExcelWorksheet nameSheet, int index)
		: base(wb, nameSheet, name, isName: true)
	{
		Name = name;
		_sheet = nameSheet;
		Index = index;
	}

	public override string ToString()
	{
		return Name;
	}
}
