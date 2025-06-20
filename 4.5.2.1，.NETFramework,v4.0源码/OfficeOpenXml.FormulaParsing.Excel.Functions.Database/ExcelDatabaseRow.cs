using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Database;

public class ExcelDatabaseRow
{
	private Dictionary<int, string> _fieldIndexes = new Dictionary<int, string>();

	private readonly Dictionary<string, object> _items = new Dictionary<string, object>();

	private int _colIndex = 1;

	public object this[string field]
	{
		get
		{
			return _items[field];
		}
		set
		{
			_items[field] = value;
			_fieldIndexes[_colIndex++] = field;
		}
	}

	public object this[int index]
	{
		get
		{
			string key = _fieldIndexes[index];
			return _items[key];
		}
	}
}
