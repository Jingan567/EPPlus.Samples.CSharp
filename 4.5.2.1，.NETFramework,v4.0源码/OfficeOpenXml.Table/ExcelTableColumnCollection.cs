using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Table;

public class ExcelTableColumnCollection : IEnumerable<ExcelTableColumn>, IEnumerable
{
	private List<ExcelTableColumn> _cols = new List<ExcelTableColumn>();

	private Dictionary<string, int> _colNames = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

	public ExcelTable Table { get; private set; }

	public int Count => _cols.Count;

	public ExcelTableColumn this[int Index]
	{
		get
		{
			if (Index < 0 || Index >= _cols.Count)
			{
				throw new ArgumentOutOfRangeException("Column index out of range");
			}
			return _cols[Index];
		}
	}

	public ExcelTableColumn this[string Name]
	{
		get
		{
			if (_colNames.ContainsKey(Name))
			{
				return _cols[_colNames[Name]];
			}
			return null;
		}
	}

	public ExcelTableColumnCollection(ExcelTable table)
	{
		Table = table;
		foreach (XmlNode item in table.TableXml.SelectNodes("//d:table/d:tableColumns/d:tableColumn", table.NameSpaceManager))
		{
			_cols.Add(new ExcelTableColumn(table.NameSpaceManager, item, table, _cols.Count));
			_colNames.Add(_cols[_cols.Count - 1].Name, _cols.Count - 1);
		}
	}

	IEnumerator<ExcelTableColumn> IEnumerable<ExcelTableColumn>.GetEnumerator()
	{
		return _cols.GetEnumerator();
	}

	IEnumerator IEnumerable.GetEnumerator()
	{
		return _cols.GetEnumerator();
	}

	internal string GetUniqueName(string name)
	{
		if (_colNames.ContainsKey(name))
		{
			string text = name;
			int num = 2;
			do
			{
				text = name + num++.ToString(CultureInfo.InvariantCulture);
			}
			while (_colNames.ContainsKey(text));
			return text;
		}
		return name;
	}
}
