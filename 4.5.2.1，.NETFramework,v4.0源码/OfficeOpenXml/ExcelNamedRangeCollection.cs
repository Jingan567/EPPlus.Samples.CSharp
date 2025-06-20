using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml;

public class ExcelNamedRangeCollection : IEnumerable<ExcelNamedRange>, IEnumerable
{
	internal ExcelWorksheet _ws;

	internal ExcelWorkbook _wb;

	private List<ExcelNamedRange> _list = new List<ExcelNamedRange>();

	private Dictionary<string, int> _dic = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

	public int Count => _dic.Count;

	public ExcelNamedRange this[string Name] => _list[_dic[Name]];

	public ExcelNamedRange this[int Index] => _list[Index];

	internal ExcelNamedRangeCollection(ExcelWorkbook wb)
	{
		_wb = wb;
		_ws = null;
	}

	internal ExcelNamedRangeCollection(ExcelWorkbook wb, ExcelWorksheet ws)
	{
		_wb = wb;
		_ws = ws;
	}

	public ExcelNamedRange Add(string Name, ExcelRangeBase Range)
	{
		ExcelNamedRange excelNamedRange = ((!Range.IsName) ? new ExcelNamedRange(Name, _ws, Range.Worksheet, Range.Address, _dic.Count) : new ExcelNamedRange(Name, _wb, _ws, _dic.Count));
		AddName(Name, excelNamedRange);
		return excelNamedRange;
	}

	private void AddName(string Name, ExcelNamedRange item)
	{
		_dic.Add(Name, _list.Count);
		_list.Add(item);
	}

	public ExcelNamedRange AddValue(string Name, object value)
	{
		ExcelNamedRange excelNamedRange = new ExcelNamedRange(Name, _wb, _ws, _dic.Count);
		excelNamedRange.NameValue = value;
		AddName(Name, excelNamedRange);
		return excelNamedRange;
	}

	[Obsolete("Call AddFormula() instead.  See Issue Tracker Id #14687")]
	public ExcelNamedRange AddFormla(string Name, string Formula)
	{
		return AddFormula(Name, Formula);
	}

	public ExcelNamedRange AddFormula(string Name, string Formula)
	{
		ExcelNamedRange excelNamedRange = new ExcelNamedRange(Name, _wb, _ws, _dic.Count);
		excelNamedRange.NameFormula = Formula;
		AddName(Name, excelNamedRange);
		return excelNamedRange;
	}

	internal void Insert(int rowFrom, int colFrom, int rows, int cols)
	{
		Insert(rowFrom, colFrom, rows, cols, (ExcelNamedRange n) => true);
	}

	internal void Insert(int rowFrom, int colFrom, int rows, int cols, Func<ExcelNamedRange, bool> filter)
	{
		foreach (ExcelNamedRange item in _list.Where(filter))
		{
			InsertRows(rowFrom, rows, item);
			InsertColumns(colFrom, cols, item);
		}
	}

	internal void Delete(int rowFrom, int colFrom, int rows, int cols)
	{
		Delete(rowFrom, colFrom, rows, cols, (ExcelNamedRange n) => true);
	}

	internal void Delete(int rowFrom, int colFrom, int rows, int cols, Func<ExcelNamedRange, bool> filter)
	{
		foreach (ExcelNamedRange item in _list.Where(filter))
		{
			ExcelAddressBase excelAddressBase = ((cols <= 0 || rowFrom != 0 || rows < 1048576) ? item.DeleteRow(rowFrom, rows) : item.DeleteColumn(colFrom, cols));
			if (excelAddressBase == null)
			{
				item.Address = "#REF!";
			}
			else
			{
				item.Address = excelAddressBase.Address;
			}
		}
	}

	private void InsertColumns(int colFrom, int cols, ExcelNamedRange namedRange)
	{
		if (colFrom > 0)
		{
			if (colFrom <= namedRange.Start.Column)
			{
				string address = ExcelCellBase.GetAddress(namedRange.Start.Row, namedRange.Start.Column + cols, namedRange.End.Row, namedRange.End.Column + cols);
				namedRange.Address = BuildNewAddress(namedRange, address);
			}
			else if (colFrom <= namedRange.End.Column && namedRange.End.Column + cols < 16384)
			{
				string address2 = ExcelCellBase.GetAddress(namedRange.Start.Row, namedRange.Start.Column, namedRange.End.Row, namedRange.End.Column + cols);
				namedRange.Address = BuildNewAddress(namedRange, address2);
			}
		}
	}

	private static string BuildNewAddress(ExcelNamedRange namedRange, string newAddress)
	{
		if (namedRange.FullAddress.Contains("!"))
		{
			newAddress = ExcelCellBase.GetFullAddress(namedRange.FullAddress.Split('!')[0].Trim('\''), newAddress);
		}
		return newAddress;
	}

	private void InsertRows(int rowFrom, int rows, ExcelNamedRange namedRange)
	{
		if (rows > 0)
		{
			if (rowFrom <= namedRange.Start.Row)
			{
				string address = ExcelCellBase.GetAddress(namedRange.Start.Row + rows, namedRange.Start.Column, namedRange.End.Row + rows, namedRange.End.Column);
				namedRange.Address = BuildNewAddress(namedRange, address);
			}
			else if (rowFrom <= namedRange.End.Row && namedRange.End.Row + rows <= 1048576)
			{
				string address2 = ExcelCellBase.GetAddress(namedRange.Start.Row, namedRange.Start.Column, namedRange.End.Row + rows, namedRange.End.Column);
				namedRange.Address = BuildNewAddress(namedRange, address2);
			}
		}
	}

	public void Remove(string Name)
	{
		if (_dic.ContainsKey(Name))
		{
			int num = _dic[Name];
			for (int i = num + 1; i < _list.Count; i++)
			{
				_dic.Remove(_list[i].Name);
				_list[i].Index--;
				_dic.Add(_list[i].Name, _list[i].Index);
			}
			_dic.Remove(Name);
			_list.RemoveAt(num);
		}
	}

	public bool ContainsKey(string key)
	{
		return _dic.ContainsKey(key);
	}

	public IEnumerator<ExcelNamedRange> GetEnumerator()
	{
		return _list.GetEnumerator();
	}

	IEnumerator IEnumerable.GetEnumerator()
	{
		return _list.GetEnumerator();
	}

	internal void Clear()
	{
		while (Count > 0)
		{
			Remove(_list[0].Name);
		}
	}
}
