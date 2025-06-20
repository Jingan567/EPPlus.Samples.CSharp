using System;
using System.Collections;
using System.Collections.Generic;

namespace OfficeOpenXml;

internal class RangeCollection : IEnumerator<IRangeID>, IDisposable, IEnumerator, IEnumerable
{
	private class IndexItem
	{
		internal ulong RangeID;

		internal int ListPointer;

		internal IndexItem(ulong cellId)
		{
			RangeID = cellId;
		}

		internal IndexItem(ulong cellId, int listPointer)
		{
			RangeID = cellId;
			ListPointer = listPointer;
		}
	}

	internal class Compare : IComparer<IndexItem>
	{
		int IComparer<IndexItem>.Compare(IndexItem x, IndexItem y)
		{
			if (x.RangeID >= y.RangeID)
			{
				if (x.RangeID <= y.RangeID)
				{
					return 0;
				}
				return 1;
			}
			return -1;
		}
	}

	private IndexItem[] _cellIndex;

	private List<IRangeID> _cells;

	private static readonly Compare _comparer = new Compare();

	private int _ix = -1;

	internal IRangeID this[ulong RangeID] => _cells[_cellIndex[IndexOf(RangeID)].ListPointer];

	internal IRangeID this[int Index] => _cells[_cellIndex[Index].ListPointer];

	internal int Count => _cells.Count;

	private int _size { get; set; }

	IRangeID IEnumerator<IRangeID>.Current
	{
		get
		{
			throw new NotImplementedException();
		}
	}

	object IEnumerator.Current => _cells[_cellIndex[_ix].ListPointer];

	internal RangeCollection(List<IRangeID> cells)
	{
		_cells = cells;
		InitSize(_cells);
		for (int i = 0; i < _cells.Count; i++)
		{
			_cellIndex[i] = new IndexItem(cells[i].RangeID, i);
		}
	}

	~RangeCollection()
	{
		_cells = null;
		_cellIndex = null;
	}

	internal void Add(IRangeID cell)
	{
		int num = IndexOf(cell.RangeID);
		if (num >= 0)
		{
			throw new Exception("Item already exists");
		}
		Insert(~num, cell);
	}

	internal void Delete(ulong key)
	{
		int num = IndexOf(key);
		if (num < 0)
		{
			throw new Exception("Key does not exists");
		}
		int listPointer = _cellIndex[num].ListPointer;
		Array.Copy(_cellIndex, num + 1, _cellIndex, num, _cells.Count - num - 1);
		_cells.RemoveAt(listPointer);
		for (int i = 0; i < _cells.Count; i++)
		{
			if (_cellIndex[i].ListPointer >= listPointer)
			{
				_cellIndex[i].ListPointer--;
			}
		}
	}

	internal int IndexOf(ulong key)
	{
		return Array.BinarySearch(_cellIndex, 0, _cells.Count, new IndexItem(key), _comparer);
	}

	internal bool ContainsKey(ulong key)
	{
		if (IndexOf(key) >= 0)
		{
			return true;
		}
		return false;
	}

	internal int InsertRowsUpdateIndex(ulong rowID, int rows)
	{
		int num = IndexOf(rowID);
		if (num < 0)
		{
			num = ~num;
		}
		ulong num2 = (ulong)((long)rows << 29);
		for (int i = num; i < _cells.Count; i++)
		{
			_cellIndex[i].RangeID += num2;
		}
		return num;
	}

	internal int InsertRows(ulong rowID, int rows)
	{
		int num = IndexOf(rowID);
		if (num < 0)
		{
			num = ~num;
		}
		ulong num2 = (ulong)((long)rows << 29);
		for (int i = num; i < _cells.Count; i++)
		{
			_cellIndex[i].RangeID += num2;
			_cells[_cellIndex[i].ListPointer].RangeID += num2;
		}
		return num;
	}

	internal int DeleteRows(ulong rowID, int rows, bool updateCells)
	{
		ulong num = (ulong)((long)rows << 29);
		int num2 = IndexOf(rowID);
		if (num2 < 0)
		{
			num2 = ~num2;
		}
		if (num2 >= _cells.Count || _cellIndex[num2] == null)
		{
			return -1;
		}
		while (num2 < _cells.Count && _cellIndex[num2].RangeID < rowID + num)
		{
			Delete(_cellIndex[num2].RangeID);
		}
		int num3 = IndexOf(rowID + num);
		if (num3 < 0)
		{
			num3 = ~num3;
		}
		for (int i = num3; i < _cells.Count; i++)
		{
			_cellIndex[i].RangeID -= num;
			if (updateCells)
			{
				_cells[_cellIndex[i].ListPointer].RangeID -= num;
			}
		}
		return num2;
	}

	internal void InsertColumn(ulong ColumnID, int columns)
	{
		throw new Exception("Working on it...");
	}

	internal void DeleteColumn(ulong ColumnID, int columns)
	{
		throw new Exception("Working on it...");
	}

	private void InitSize(List<IRangeID> _cells)
	{
		_size = 128;
		while (_cells.Count > _size)
		{
			_size <<= 1;
		}
		_cellIndex = new IndexItem[_size];
	}

	private void CheckSize()
	{
		if (_cells.Count >= _size)
		{
			_size <<= 1;
			Array.Resize(ref _cellIndex, _size);
		}
	}

	private void Insert(int ix, IRangeID cell)
	{
		CheckSize();
		Array.Copy(_cellIndex, ix, _cellIndex, ix + 1, _cells.Count - ix);
		_cellIndex[ix] = new IndexItem(cell.RangeID, _cells.Count);
		_cells.Add(cell);
	}

	void IDisposable.Dispose()
	{
		_ix = -1;
	}

	bool IEnumerator.MoveNext()
	{
		_ix++;
		return _ix < _cells.Count;
	}

	void IEnumerator.Reset()
	{
		_ix = -1;
	}

	IEnumerator IEnumerable.GetEnumerator()
	{
		return MemberwiseClone() as IEnumerator;
	}
}
