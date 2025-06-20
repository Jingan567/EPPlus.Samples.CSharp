using System;
using System.Collections;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Table;

public class ExcelTableCollection : IEnumerable<ExcelTable>, IEnumerable
{
	private List<ExcelTable> _tables = new List<ExcelTable>();

	internal Dictionary<string, int> _tableNames = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

	private ExcelWorksheet _ws;

	public int Count => _tables.Count;

	public ExcelTable this[int Index]
	{
		get
		{
			if (Index < 0 || Index >= _tables.Count)
			{
				throw new ArgumentOutOfRangeException("Table index out of range");
			}
			return _tables[Index];
		}
	}

	public ExcelTable this[string Name]
	{
		get
		{
			if (_tableNames.ContainsKey(Name))
			{
				return _tables[_tableNames[Name]];
			}
			return null;
		}
	}

	internal ExcelTableCollection(ExcelWorksheet ws)
	{
		_ = ws._package.Package;
		_ws = ws;
		foreach (XmlElement item in ws.WorksheetXml.SelectNodes("//d:tableParts/d:tablePart", ws.NameSpaceManager))
		{
			ExcelTable excelTable = new ExcelTable(ws.Part.GetRelationship(item.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")), ws);
			_tableNames.Add(excelTable.Name, _tables.Count);
			_tables.Add(excelTable);
		}
	}

	private ExcelTable Add(ExcelTable tbl)
	{
		_tables.Add(tbl);
		_tableNames.Add(tbl.Name, _tables.Count - 1);
		if (tbl.Id >= _ws.Workbook._nextTableID)
		{
			_ws.Workbook._nextTableID = tbl.Id + 1;
		}
		return tbl;
	}

	public ExcelTable Add(ExcelAddressBase Range, string Name)
	{
		if (Range.WorkSheet != null && Range.WorkSheet != _ws.Name)
		{
			throw new ArgumentException("Range does not belong to worksheet", "Range");
		}
		if (string.IsNullOrEmpty(Name))
		{
			Name = GetNewTableName();
		}
		else if (_ws.Workbook.ExistsTableName(Name))
		{
			throw new ArgumentException("Tablename is not unique");
		}
		ValidateTableName(Name);
		foreach (ExcelTable table in _tables)
		{
			if (table.Address.Collide(Range) != 0)
			{
				throw new ArgumentException($"Table range collides with table {table.Name}");
			}
		}
		return Add(new ExcelTable(_ws, Range, Name, _ws.Workbook._nextTableID));
	}

	private void ValidateTableName(string Name)
	{
		if (string.IsNullOrEmpty(Name))
		{
			throw new ArgumentException("Tablename is null or empty");
		}
		char c = Name[0];
		if (!char.IsLetter(c) && c != '_' && c != '\\')
		{
			throw new ArgumentException("Tablename start with invalid character");
		}
		if (Name.Contains(" "))
		{
			throw new ArgumentException("Tablename has spaces");
		}
	}

	public void Delete(int Index, bool ClearRange = false)
	{
		Delete(this[Index], ClearRange);
	}

	public void Delete(string Name, bool ClearRange = false)
	{
		if (this[Name] == null)
		{
			throw new ArgumentOutOfRangeException($"Cannot delete non-existant table {Name} in sheet {_ws.Name}.");
		}
		Delete(this[Name], ClearRange);
	}

	public void Delete(ExcelTable Table, bool ClearRange = false)
	{
		if (!_tables.Contains(Table))
		{
			throw new ArgumentOutOfRangeException("Table", $"Table {Table.Name} does not exist in this collection");
		}
		lock (this)
		{
			ExcelRange excelRange = _ws.Cells[Table.Address.Address];
			_tableNames.Remove(Table.Name);
			_tables.Remove(Table);
			foreach (ExcelWorksheet worksheet in Table.WorkSheet.Workbook.Worksheets)
			{
				foreach (ExcelTable table in worksheet.Tables)
				{
					if (table.Id > Table.Id)
					{
						table.Id--;
					}
				}
				Table.WorkSheet.Workbook._nextTableID--;
			}
			if (ClearRange)
			{
				excelRange.Clear();
			}
		}
	}

	internal string GetNewTableName()
	{
		string text = "Table1";
		int num = 2;
		while (_ws.Workbook.ExistsTableName(text))
		{
			text = $"Table{num++}";
		}
		return text;
	}

	public ExcelTable GetFromRange(ExcelRangeBase Range)
	{
		foreach (ExcelTable table in Range.Worksheet.Tables)
		{
			if (table.Address._address == Range._address)
			{
				return table;
			}
		}
		return null;
	}

	public IEnumerator<ExcelTable> GetEnumerator()
	{
		return _tables.GetEnumerator();
	}

	IEnumerator IEnumerable.GetEnumerator()
	{
		return _tables.GetEnumerator();
	}
}
