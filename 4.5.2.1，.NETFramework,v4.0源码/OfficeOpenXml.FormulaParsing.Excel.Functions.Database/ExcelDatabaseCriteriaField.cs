namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Database;

public class ExcelDatabaseCriteriaField
{
	public string FieldName { get; private set; }

	public int? FieldIndex { get; private set; }

	public ExcelDatabaseCriteriaField(string fieldName)
	{
		FieldName = fieldName;
	}

	public ExcelDatabaseCriteriaField(int fieldIndex)
	{
		FieldIndex = fieldIndex;
	}

	public override string ToString()
	{
		if (!string.IsNullOrEmpty(FieldName))
		{
			return FieldName;
		}
		return base.ToString();
	}
}
