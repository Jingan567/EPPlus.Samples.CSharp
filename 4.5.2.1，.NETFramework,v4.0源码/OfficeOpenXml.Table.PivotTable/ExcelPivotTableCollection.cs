using System;
using System.Collections;
using System.Collections.Generic;
using OfficeOpenXml.Packaging;

namespace OfficeOpenXml.Table.PivotTable;

public class ExcelPivotTableCollection : IEnumerable<ExcelPivotTable>, IEnumerable
{
	private List<ExcelPivotTable> _pivotTables = new List<ExcelPivotTable>();

	internal Dictionary<string, int> _pivotTableNames = new Dictionary<string, int>();

	private ExcelWorksheet _ws;

	public int Count => _pivotTables.Count;

	public ExcelPivotTable this[int Index]
	{
		get
		{
			if (Index < 0 || Index >= _pivotTables.Count)
			{
				throw new ArgumentOutOfRangeException("PivotTable index out of range");
			}
			return _pivotTables[Index];
		}
	}

	public ExcelPivotTable this[string Name]
	{
		get
		{
			if (_pivotTableNames.ContainsKey(Name))
			{
				return _pivotTables[_pivotTableNames[Name]];
			}
			return null;
		}
	}

	internal ExcelPivotTableCollection(ExcelWorksheet ws)
	{
		_ = ws._package.Package;
		_ws = ws;
		foreach (ZipPackageRelationship relationship in ws.Part.GetRelationships())
		{
			if (relationship.RelationshipType == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable")
			{
				ExcelPivotTable excelPivotTable = new ExcelPivotTable(relationship, ws);
				_pivotTableNames.Add(excelPivotTable.Name, _pivotTables.Count);
				_pivotTables.Add(excelPivotTable);
			}
		}
	}

	private ExcelPivotTable Add(ExcelPivotTable tbl)
	{
		_pivotTables.Add(tbl);
		_pivotTableNames.Add(tbl.Name, _pivotTables.Count - 1);
		if (tbl.CacheID >= _ws.Workbook._nextPivotTableID)
		{
			_ws.Workbook._nextPivotTableID = tbl.CacheID + 1;
		}
		return tbl;
	}

	public ExcelPivotTable Add(ExcelAddressBase Range, ExcelRangeBase Source, string Name)
	{
		if (string.IsNullOrEmpty(Name))
		{
			Name = GetNewTableName();
		}
		if (Range.WorkSheet != _ws.Name)
		{
			throw new Exception("The Range must be in the current worksheet");
		}
		if (_ws.Workbook.ExistsTableName(Name))
		{
			throw new ArgumentException("Tablename is not unique");
		}
		foreach (ExcelPivotTable pivotTable in _pivotTables)
		{
			if (pivotTable.Address.Collide(Range) != 0)
			{
				throw new ArgumentException($"Table range collides with table {pivotTable.Name}");
			}
		}
		return Add(new ExcelPivotTable(_ws, Range, Source, Name, _ws.Workbook._nextPivotTableID++));
	}

	internal string GetNewTableName()
	{
		string text = "Pivottable1";
		int num = 2;
		while (_ws.Workbook.ExistsPivotTableName(text))
		{
			text = $"Pivottable{num++}";
		}
		return text;
	}

	public IEnumerator<ExcelPivotTable> GetEnumerator()
	{
		return _pivotTables.GetEnumerator();
	}

	IEnumerator IEnumerable.GetEnumerator()
	{
		return _pivotTables.GetEnumerator();
	}
}
