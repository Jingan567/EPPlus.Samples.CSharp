namespace OfficeOpenXml.FormulaParsing;

public class EpplusNameValueProvider : INameValueProvider
{
	private ExcelDataProvider _excelDataProvider;

	private ExcelNamedRangeCollection _values;

	public EpplusNameValueProvider(ExcelDataProvider excelDataProvider)
	{
		_excelDataProvider = excelDataProvider;
		_values = _excelDataProvider.GetWorkbookNameValues();
	}

	public virtual bool IsNamedValue(string key, string ws)
	{
		if (ws != null)
		{
			ExcelNamedRangeCollection worksheetNames = _excelDataProvider.GetWorksheetNames(ws);
			if (worksheetNames != null && worksheetNames.ContainsKey(key))
			{
				return true;
			}
		}
		if (_values != null)
		{
			return _values.ContainsKey(key);
		}
		return false;
	}

	public virtual object GetNamedValue(string key)
	{
		return _values[key];
	}

	public virtual void Reload()
	{
		_values = _excelDataProvider.GetWorkbookNameValues();
	}
}
