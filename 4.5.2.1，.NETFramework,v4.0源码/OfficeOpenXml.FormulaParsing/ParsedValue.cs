namespace OfficeOpenXml.FormulaParsing;

public class ParsedValue
{
	public object Value { get; private set; }

	public int RowIndex { get; private set; }

	public int ColIndex { get; private set; }

	public ParsedValue(object val, int rowIndex, int colIndex)
	{
		Value = val;
		RowIndex = rowIndex;
		ColIndex = colIndex;
	}
}
