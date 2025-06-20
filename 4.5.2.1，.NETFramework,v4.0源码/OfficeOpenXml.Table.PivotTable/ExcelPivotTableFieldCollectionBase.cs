using System;
using System.Collections;
using System.Collections.Generic;

namespace OfficeOpenXml.Table.PivotTable;

public class ExcelPivotTableFieldCollectionBase<T> : IEnumerable<T>, IEnumerable
{
	protected ExcelPivotTable _table;

	internal List<T> _list = new List<T>();

	public int Count => _list.Count;

	public T this[int Index]
	{
		get
		{
			if (Index < 0 || Index >= _list.Count)
			{
				throw new ArgumentOutOfRangeException("Index out of range");
			}
			return _list[Index];
		}
	}

	internal ExcelPivotTableFieldCollectionBase(ExcelPivotTable table)
	{
		_table = table;
	}

	public IEnumerator<T> GetEnumerator()
	{
		return _list.GetEnumerator();
	}

	IEnumerator IEnumerable.GetEnumerator()
	{
		return _list.GetEnumerator();
	}

	internal void AddInternal(T field)
	{
		_list.Add(field);
	}

	internal void Clear()
	{
		_list.Clear();
	}
}
