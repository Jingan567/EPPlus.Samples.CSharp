using System;
using System.Collections;
using System.Collections.Generic;
using OfficeOpenXml;

internal class CellsStoreEnumerator<T> : IEnumerable<T>, IEnumerable, IEnumerator<T>, IDisposable, IEnumerator
{
	private CellStore<T> _cellStore;

	private int row;

	private int colPos;

	private int[] pagePos;

	private int[] cellPos;

	private int _startRow;

	private int _startCol;

	private int _endRow;

	private int _endCol;

	private int minRow;

	private int minColPos;

	private int maxRow;

	private int maxColPos;

	internal int Row => row;

	internal int Column
	{
		get
		{
			if (colPos == -1)
			{
				MoveNext();
			}
			if (colPos == -1)
			{
				return 0;
			}
			return _cellStore._columnIndex[colPos].Index;
		}
	}

	internal T Value
	{
		get
		{
			lock (_cellStore)
			{
				return _cellStore.GetValue(row, Column);
			}
		}
		set
		{
			lock (_cellStore)
			{
				_cellStore.SetValue(row, Column, value);
			}
		}
	}

	public string CellAddress => ExcelCellBase.GetAddress(Row, Column);

	public T Current => Value;

	object IEnumerator.Current
	{
		get
		{
			Reset();
			return this;
		}
	}

	public CellsStoreEnumerator(CellStore<T> cellStore)
		: this(cellStore, 0, 0, 1048576, 16384)
	{
	}

	public CellsStoreEnumerator(CellStore<T> cellStore, int StartRow, int StartCol, int EndRow, int EndCol)
	{
		_cellStore = cellStore;
		_startRow = StartRow;
		_startCol = StartCol;
		_endRow = EndRow;
		_endCol = EndCol;
		Init();
	}

	internal void Init()
	{
		minRow = _startRow;
		maxRow = _endRow;
		minColPos = _cellStore.GetPosition(_startCol);
		if (minColPos < 0)
		{
			minColPos = ~minColPos;
		}
		maxColPos = _cellStore.GetPosition(_endCol);
		if (maxColPos < 0)
		{
			maxColPos = ~maxColPos - 1;
		}
		row = minRow;
		colPos = minColPos - 1;
		int num = maxColPos - minColPos + 1;
		pagePos = new int[num];
		cellPos = new int[num];
		for (int i = 0; i < num; i++)
		{
			pagePos[i] = -1;
			cellPos[i] = -1;
		}
	}

	internal bool Next()
	{
		return _cellStore.GetNextCell(ref row, ref colPos, minColPos, maxRow, maxColPos);
	}

	internal bool Previous()
	{
		lock (_cellStore)
		{
			return _cellStore.GetPrevCell(ref row, ref colPos, minRow, minColPos, maxColPos);
		}
	}

	public IEnumerator<T> GetEnumerator()
	{
		Reset();
		return this;
	}

	IEnumerator IEnumerable.GetEnumerator()
	{
		Reset();
		return this;
	}

	public void Dispose()
	{
	}

	public bool MoveNext()
	{
		return Next();
	}

	public void Reset()
	{
		Init();
	}
}
