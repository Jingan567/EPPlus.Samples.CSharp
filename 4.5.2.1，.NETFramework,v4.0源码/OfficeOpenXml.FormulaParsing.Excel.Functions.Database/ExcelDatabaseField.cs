namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Database;

public class ExcelDatabaseField
{
	public string FieldName { get; private set; }

	public int ColIndex { get; private set; }

	public ExcelDatabaseField(string fieldName, int colIndex)
	{
		FieldName = fieldName;
		ColIndex = colIndex;
	}
}
