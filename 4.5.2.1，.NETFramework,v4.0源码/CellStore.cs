using System;
using System.Collections.Generic;

internal class CellStore<T> : IDisposable
{
	internal delegate void SetRangeValueDelegate(List<T> list, int index, int row, int column, object value);

	internal delegate void SetValueDelegate(List<T> list, int index, object value);

	internal const int pageBits = 10;

	internal const int PageSize = 1024;

	internal const int PageSizeMin = 1024;

	internal const int PageSizeMax = 2048;

	internal const int ColSizeMin = 32;

	internal const int PagesPerColumnMin = 32;

	private List<T> _values = new List<T>();

	internal ColumnIndex[] _columnIndex;

	internal IndexBase _searchIx = new IndexBase();

	internal IndexItem _searchItem;

	internal int ColumnCount;

	private int _colPos = -1;

	private int _row;

	internal int Count
	{
		get
		{
			int num = 0;
			for (int i = 0; i < ColumnCount; i++)
			{
				for (int j = 0; j < _columnIndex[i].PageCount; j++)
				{
					num += _columnIndex[i]._pages[j].RowCount;
				}
			}
			return num;
		}
	}

	public ulong Current => (ulong)(((long)_row << 32) | _columnIndex[_colPos].Index);

	public CellStore()
	{
		_columnIndex = new ColumnIndex[32];
	}

	~CellStore()
	{
		if (_values != null)
		{
			_values.Clear();
			_values = null;
		}
		_columnIndex = null;
	}

	internal int GetPosition(int Column)
	{
		if (Column < ColumnCount && _columnIndex[Column].Index == Column)
		{
			return Column;
		}
		_searchIx.Index = (short)Column;
		return Array.BinarySearch(_columnIndex, 0, ColumnCount, _searchIx);
	}

	internal CellStore<T> Clone()
	{
		CellStore<T> cellStore = new CellStore<T>();
		for (int i = 0; i < ColumnCount; i++)
		{
			int index = _columnIndex[i].Index;
			for (int j = 0; j < _columnIndex[i].PageCount; j++)
			{
				for (int k = 0; k < _columnIndex[i]._pages[j].RowCount; k++)
				{
					int row = _columnIndex[i]._pages[j].IndexOffset + _columnIndex[i]._pages[j].Rows[k].Index;
					cellStore.SetValue(row, index, _values[_columnIndex[i]._pages[j].Rows[k].IndexPointer]);
				}
			}
		}
		return cellStore;
	}

	internal bool GetDimension(out int fromRow, out int fromCol, out int toRow, out int toCol)
	{
		if (ColumnCount == 0)
		{
			fromRow = (fromCol = (toRow = (toCol = 0)));
			return false;
		}
		fromCol = _columnIndex[0].Index;
		int num = 0;
		if (fromCol <= 0 && ColumnCount > 1)
		{
			fromCol = _columnIndex[1].Index;
			num = 1;
		}
		else if (ColumnCount == 1 && fromCol <= 0)
		{
			fromRow = (fromCol = (toRow = (toCol = 0)));
			return false;
		}
		int num2 = ColumnCount - 1;
		while (num2 > 0 && _columnIndex[num2].PageCount != 0 && _columnIndex[num2]._pages[0].RowCount <= 1 && _columnIndex[num2]._pages[0].Rows[0].Index <= 0)
		{
			num2--;
		}
		toCol = _columnIndex[num2].Index;
		if (toCol == 0)
		{
			fromRow = (fromCol = (toRow = (toCol = 0)));
			return false;
		}
		fromRow = (toRow = 0);
		for (int i = num; i < ColumnCount; i++)
		{
			if (_columnIndex[i].PageCount != 0)
			{
				int num3 = ((_columnIndex[i]._pages[0].RowCount > 0 && _columnIndex[i]._pages[0].Rows[0].Index > 0) ? (_columnIndex[i]._pages[0].IndexOffset + _columnIndex[i]._pages[0].Rows[0].Index) : ((_columnIndex[i]._pages[0].RowCount > 1) ? (_columnIndex[i]._pages[0].IndexOffset + _columnIndex[i]._pages[0].Rows[1].Index) : ((_columnIndex[i].PageCount > 1) ? (_columnIndex[i]._pages[0].IndexOffset + _columnIndex[i]._pages[1].Rows[0].Index) : 0)));
				int num4 = _columnIndex[i].PageCount - 1;
				while (_columnIndex[i]._pages[num4].RowCount == 0 && num4 != 0)
				{
					num4--;
				}
				PageIndex pageIndex = _columnIndex[i]._pages[num4];
				int num5 = ((pageIndex.RowCount <= 0) ? num3 : (pageIndex.IndexOffset + pageIndex.Rows[pageIndex.RowCount - 1].Index));
				if (num3 > 0 && (num3 < fromRow || fromRow == 0))
				{
					fromRow = num3;
				}
				if (num3 > 0 && (num5 > toRow || toRow == 0))
				{
					toRow = num5;
				}
			}
		}
		if (fromRow <= 0 || toRow <= 0)
		{
			fromRow = (fromCol = (toRow = (toCol = 0)));
			return false;
		}
		return true;
	}

	internal int FindNext(int Column)
	{
		int position = GetPosition(Column);
		if (position < 0)
		{
			return ~position;
		}
		return position;
	}

	internal T GetValue(int Row, int Column)
	{
		int pointer = GetPointer(Row, Column);
		if (pointer >= 0)
		{
			return _values[pointer];
		}
		return default(T);
	}

	private int GetPointer(int Row, int Column)
	{
		int position = GetPosition(Column);
		if (position >= 0)
		{
			int position2 = _columnIndex[position].GetPosition(Row);
			if (position2 >= 0 && position2 < _columnIndex[position].PageCount)
			{
				PageIndex pageIndex = _columnIndex[position]._pages[position2];
				if (pageIndex.MinIndex > Row)
				{
					position2--;
					if (position2 < 0)
					{
						return -1;
					}
					pageIndex = _columnIndex[position]._pages[position2];
				}
				short index = (short)(Row - pageIndex.IndexOffset);
				_searchItem.Index = index;
				int num = Array.BinarySearch(pageIndex.Rows, 0, pageIndex.RowCount, _searchItem);
				if (num >= 0)
				{
					return pageIndex.Rows[num].IndexPointer;
				}
				return -1;
			}
			return -1;
		}
		return -1;
	}

	internal bool Exists(int Row, int Column)
	{
		return GetPointer(Row, Column) >= 0;
	}

	internal bool Exists(int Row, int Column, ref T value)
	{
		int pointer = GetPointer(Row, Column);
		if (pointer >= 0)
		{
			value = _values[pointer];
			return true;
		}
		return false;
	}

	internal void SetValue(int Row, int Column, T Value)
	{
		lock (_columnIndex)
		{
			int position = GetPosition(Column);
			short num = (short)(Row >> 10);
			if (position >= 0)
			{
				int num2 = _columnIndex[position].GetPosition(Row);
				if (num2 < 0)
				{
					num2 = ~num2;
					if (num2 - 1 < 0 || _columnIndex[position]._pages[num2 - 1].IndexOffset + 1024 - 1 < Row)
					{
						AddPage(_columnIndex[position], num2, num);
					}
					else
					{
						num2--;
					}
				}
				if (num2 >= _columnIndex[position].PageCount)
				{
					AddPage(_columnIndex[position], num2, num);
				}
				PageIndex pageIndex = _columnIndex[position]._pages[num2];
				if ((pageIndex.MinIndex > Row || pageIndex.MaxIndex < Row) && pageIndex.IndexExpanded > Row)
				{
					num2--;
					num--;
					if (num2 < 0)
					{
						throw new Exception("Unexpected error when setting value");
					}
					pageIndex = _columnIndex[position]._pages[num2];
				}
				short num3 = (short)(Row - ((pageIndex.Index << 10) + pageIndex.Offset));
				_searchItem.Index = num3;
				int num4 = Array.BinarySearch(pageIndex.Rows, 0, pageIndex.RowCount, _searchItem);
				if (num4 < 0)
				{
					num4 = ~num4;
					AddCell(_columnIndex[position], num2, num4, num3, Value);
				}
				else
				{
					_values[pageIndex.Rows[num4].IndexPointer] = Value;
				}
			}
			else
			{
				position = ~position;
				AddColumn(position, Column);
				AddPage(_columnIndex[position], 0, num);
				short ix = (short)(Row - (num << 10));
				AddCell(_columnIndex[position], 0, 0, ix, Value);
			}
		}
	}

	internal void SetRangeValueSpecial(int fromRow, int fromColumn, int toRow, int toColumn, SetRangeValueDelegate Updater, object Value)
	{
		lock (_columnIndex)
		{
			Dictionary<short, List<int>> dictionary = new Dictionary<short, List<int>>();
			for (int i = fromRow; i <= toRow; i++)
			{
				short key = (short)(i >> 10);
				if (!dictionary.ContainsKey(key))
				{
					dictionary.Add(key, new List<int>());
				}
				dictionary[key].Add(i);
			}
			for (int j = fromColumn; j <= toColumn; j++)
			{
				int num = GetPosition(j);
				foreach (KeyValuePair<short, List<int>> item in dictionary)
				{
					short num2 = item.Key;
					foreach (int item2 in item.Value)
					{
						if (num >= 0)
						{
							int num3 = _columnIndex[num].GetPosition(item2);
							if (num3 < 0)
							{
								num3 = ~num3;
								if (num3 - 1 < 0 || _columnIndex[num]._pages[num3 - 1].IndexOffset + 1024 - 1 < item2)
								{
									AddPage(_columnIndex[num], num3, num2);
								}
								else
								{
									num3--;
								}
							}
							if (num3 >= _columnIndex[num].PageCount)
							{
								AddPage(_columnIndex[num], num3, num2);
							}
							PageIndex pageIndex = _columnIndex[num]._pages[num3];
							if (pageIndex.IndexOffset > item2)
							{
								num3--;
								num2--;
								if (num3 < 0)
								{
									throw new Exception("Unexpected error when setting value");
								}
								pageIndex = _columnIndex[num]._pages[num3];
							}
							short num4 = (short)(item2 - ((pageIndex.Index << 10) + pageIndex.Offset));
							_searchItem.Index = num4;
							int num5 = Array.BinarySearch(pageIndex.Rows, 0, pageIndex.RowCount, _searchItem);
							if (num5 < 0)
							{
								num5 = ~num5;
								AddCell(_columnIndex[num], num3, num5, num4, default(T));
								Updater(_values, pageIndex.Rows[num5].IndexPointer, item2, j, Value);
							}
							else
							{
								Updater(_values, pageIndex.Rows[num5].IndexPointer, item2, j, Value);
							}
						}
						else
						{
							num = ~num;
							AddColumn(num, j);
							AddPage(_columnIndex[num], 0, num2);
							short ix = (short)(item2 - (num2 << 10));
							AddCell(_columnIndex[num], 0, 0, ix, default(T));
							Updater(_values, _columnIndex[num]._pages[0].Rows[0].IndexPointer, item2, j, Value);
						}
					}
				}
			}
		}
	}

	internal void SetValueSpecial(int Row, int Column, SetValueDelegate Updater, object Value)
	{
		lock (_columnIndex)
		{
			int position = GetPosition(Column);
			short num = (short)(Row >> 10);
			if (position >= 0)
			{
				int num2 = _columnIndex[position].GetPosition(Row);
				if (num2 < 0)
				{
					num2 = ~num2;
					if (num2 - 1 < 0 || _columnIndex[position]._pages[num2 - 1].IndexOffset + 1024 - 1 < Row)
					{
						AddPage(_columnIndex[position], num2, num);
					}
					else
					{
						num2--;
					}
				}
				if (num2 >= _columnIndex[position].PageCount)
				{
					AddPage(_columnIndex[position], num2, num);
				}
				PageIndex pageIndex = _columnIndex[position]._pages[num2];
				if (pageIndex.IndexOffset > Row)
				{
					num2--;
					num--;
					if (num2 < 0)
					{
						throw new Exception("Unexpected error when setting value");
					}
					pageIndex = _columnIndex[position]._pages[num2];
				}
				short num3 = (short)(Row - ((pageIndex.Index << 10) + pageIndex.Offset));
				_searchItem.Index = num3;
				int num4 = Array.BinarySearch(pageIndex.Rows, 0, pageIndex.RowCount, _searchItem);
				if (num4 < 0)
				{
					num4 = ~num4;
					AddCell(_columnIndex[position], num2, num4, num3, default(T));
					Updater(_values, pageIndex.Rows[num4].IndexPointer, Value);
				}
				else
				{
					Updater(_values, pageIndex.Rows[num4].IndexPointer, Value);
				}
			}
			else
			{
				position = ~position;
				AddColumn(position, Column);
				AddPage(_columnIndex[position], 0, num);
				short ix = (short)(Row - (num << 10));
				AddCell(_columnIndex[position], 0, 0, ix, default(T));
				Updater(_values, _columnIndex[position]._pages[0].Rows[0].IndexPointer, Value);
			}
		}
	}

	internal void Insert(int fromRow, int fromCol, int rows, int columns)
	{
		lock (_columnIndex)
		{
			if (columns > 0)
			{
				int num = GetPosition(fromCol);
				if (num < 0)
				{
					num = ~num;
				}
				for (int i = num; i < ColumnCount; i++)
				{
					_columnIndex[i].Index += (short)columns;
				}
				return;
			}
			int num2 = fromRow >> 10;
			for (int j = 0; j < ColumnCount; j++)
			{
				ColumnIndex columnIndex = _columnIndex[j];
				int position = columnIndex.GetPosition(fromRow);
				if (position >= 0)
				{
					if (fromRow >= columnIndex._pages[position].MinIndex && fromRow <= columnIndex._pages[position].MaxIndex)
					{
						int offset = fromRow - columnIndex._pages[position].IndexOffset;
						int num3 = columnIndex._pages[position].GetPosition(offset);
						if (num3 < 0)
						{
							num3 = ~num3;
						}
						UpdateIndexOffset(columnIndex, position, num3, fromRow, rows);
					}
					else if (columnIndex._pages[position].MinIndex > fromRow - 1 && position > 0)
					{
						int offset2 = fromRow - (num2 - 1 << 10);
						int position2 = columnIndex._pages[position - 1].GetPosition(offset2);
						if (position2 > 0 && position > 0)
						{
							UpdateIndexOffset(columnIndex, position - 1, position2, fromRow, rows);
						}
					}
					else if (columnIndex.PageCount >= position + 1)
					{
						int offset3 = fromRow - columnIndex._pages[position].IndexOffset;
						int num4 = columnIndex._pages[position].GetPosition(offset3);
						if (num4 < 0)
						{
							num4 = ~num4;
						}
						if (columnIndex._pages[position].RowCount > num4)
						{
							UpdateIndexOffset(columnIndex, position, num4, fromRow, rows);
						}
						else
						{
							UpdateIndexOffset(columnIndex, position + 1, 0, fromRow, rows);
						}
					}
				}
				else
				{
					UpdateIndexOffset(columnIndex, ~position, 0, fromRow, rows);
				}
			}
		}
	}

	internal void Clear(int fromRow, int fromCol, int rows, int columns)
	{
		Delete(fromRow, fromCol, rows, columns, shift: false);
	}

	internal void Delete(int fromRow, int fromCol, int rows, int columns)
	{
		Delete(fromRow, fromCol, rows, columns, shift: true);
	}

	internal void Delete(int fromRow, int fromCol, int rows, int columns, bool shift)
	{
		lock (_columnIndex)
		{
			if (columns > 0 && fromRow == 0 && rows >= 1048576)
			{
				DeleteColumns(fromCol, columns, shift);
				return;
			}
			int num = fromCol + columns - 1;
			int num2 = fromRow >> 10;
			for (int i = 0; i < ColumnCount; i++)
			{
				ColumnIndex columnIndex = _columnIndex[i];
				if (columnIndex.Index < fromCol)
				{
					continue;
				}
				if (columnIndex.Index > num)
				{
					break;
				}
				int num3 = columnIndex.GetPosition(fromRow);
				if (num3 < 0)
				{
					num3 = ~num3;
				}
				if (num3 >= columnIndex.PageCount)
				{
					continue;
				}
				PageIndex pageIndex = columnIndex._pages[num3];
				if (shift && pageIndex.RowCount > 0 && pageIndex.MinIndex > fromRow && pageIndex.MaxIndex >= fromRow + rows)
				{
					int num4 = pageIndex.MinIndex - fromRow;
					if (num4 >= rows)
					{
						pageIndex.Offset -= rows;
						UpdatePageOffset(columnIndex, num3, rows);
						continue;
					}
					rows -= num4;
					pageIndex.Offset -= num4;
					UpdatePageOffset(columnIndex, num3, num4);
				}
				if (pageIndex.RowCount > 0 && pageIndex.MinIndex <= fromRow + rows - 1 && pageIndex.MaxIndex >= fromRow)
				{
					int num5 = fromRow + rows;
					int num6 = DeleteCells(columnIndex._pages[num3], fromRow, num5, shift);
					if (shift && num6 != fromRow)
					{
						UpdatePageOffset(columnIndex, num3, num6 - fromRow);
					}
					if (num5 <= num6 || num3 >= columnIndex.PageCount || columnIndex._pages[num3].MinIndex >= num5)
					{
						continue;
					}
					num3 = ((num6 == fromRow) ? num3 : (num3 + 1));
					int num7 = DeletePage(shift ? fromRow : num6, num5 - num6, columnIndex, num3, shift);
					if (num7 > 0)
					{
						int num8 = (shift ? fromRow : (num5 - num7));
						num3 = columnIndex.GetPosition(num8);
						num6 = DeleteCells(columnIndex._pages[num3], num8, shift ? (num8 + num7) : num5, shift);
						if (shift)
						{
							UpdatePageOffset(columnIndex, num3, num7);
						}
					}
				}
				else if (num3 > 0 && columnIndex._pages[num3].IndexOffset > fromRow)
				{
					int offset = fromRow + rows - 1 - (num2 - 1 << 10);
					int position = columnIndex._pages[num3 - 1].GetPosition(offset);
					if (position > 0 && num3 > 0 && shift)
					{
						UpdateIndexOffset(columnIndex, num3 - 1, position, fromRow + rows - 1, -rows);
					}
				}
				else if (shift && num3 + 1 < columnIndex.PageCount)
				{
					UpdateIndexOffset(columnIndex, num3 + 1, 0, columnIndex._pages[num3 + 1].MinIndex, -rows);
				}
			}
		}
	}

	private void UpdatePageOffset(ColumnIndex column, int pagePos, int rows)
	{
		if (++pagePos >= column.PageCount)
		{
			return;
		}
		for (int i = pagePos; i < column.PageCount; i++)
		{
			if (column._pages[i].Offset - rows <= -1024)
			{
				column._pages[i].Index--;
				column._pages[i].Offset -= rows - 1024;
			}
			else
			{
				column._pages[i].Offset -= rows;
			}
		}
		if (Math.Abs(column._pages[pagePos].Offset) > 1024 || Math.Abs(column._pages[pagePos].Rows[column._pages[pagePos].RowCount - 1].Index) > 2048)
		{
			rows = ResetPageOffset(column, pagePos, rows);
		}
	}

	private int ResetPageOffset(ColumnIndex column, int pagePos, int rows)
	{
		PageIndex pageIndex = column._pages[pagePos];
		short num = 0;
		if (pageIndex.Offset < -1024)
		{
			PageIndex pageIndex2 = column._pages[pagePos - 1];
			num = -1;
			if (pageIndex.Index - 1 == pageIndex2.Index)
			{
				if (pageIndex.IndexOffset + pageIndex.Rows[pageIndex.RowCount - 1].Index - pageIndex2.IndexOffset + pageIndex2.Rows[0].Index <= 2048)
				{
					MergePage(column, pagePos - 1);
				}
			}
			else
			{
				pageIndex.Index -= num;
				pageIndex.Offset += 1024;
			}
		}
		else if (pageIndex.Offset > 1024)
		{
			PageIndex pageIndex2 = column._pages[pagePos + 1];
			num = 1;
			if (pageIndex.Index + 1 != pageIndex2.Index)
			{
				pageIndex.Index += num;
				pageIndex.Offset += 1024;
			}
		}
		return rows;
	}

	private int DeletePage(int fromRow, int rows, ColumnIndex column, int pagePos, bool shift)
	{
		PageIndex pageIndex = column._pages[pagePos];
		int num = rows;
		while (pageIndex != null && pageIndex.MinIndex >= fromRow && ((shift && pageIndex.MaxIndex < fromRow + rows) || (!shift && pageIndex.MaxIndex < fromRow + num)))
		{
			int num2 = pageIndex.MaxIndex - pageIndex.MinIndex + 1;
			rows -= num2;
			_ = pageIndex.Offset;
			Array.Copy(column._pages, pagePos + 1, column._pages, pagePos, column.PageCount - pagePos + 1);
			column.PageCount--;
			if (column.PageCount == 0)
			{
				return 0;
			}
			if (shift)
			{
				for (int i = pagePos; i < column.PageCount; i++)
				{
					column._pages[i].Offset -= num2;
					if (column._pages[i].Offset <= -1024)
					{
						column._pages[i].Index--;
						column._pages[i].Offset += 1024;
					}
				}
			}
			if (column.PageCount > pagePos)
			{
				pageIndex = column._pages[pagePos];
				continue;
			}
			return 0;
		}
		return rows;
	}

	private int DeleteCells(PageIndex page, int fromRow, int toRow, bool shift)
	{
		int num = page.GetPosition(fromRow - page.IndexOffset);
		if (num < 0)
		{
			num = ~num;
		}
		int maxIndex = page.MaxIndex;
		int num2 = toRow - page.IndexOffset;
		if (num2 > 2048)
		{
			num2 = 2048;
		}
		int num3 = page.GetPosition(num2);
		if (num3 < 0)
		{
			num3 = ~num3;
		}
		if (num <= num3 && num < page.RowCount && page.GetIndex(num) < toRow)
		{
			if (toRow > page.MaxIndex)
			{
				if (fromRow == page.MinIndex)
				{
					return fromRow;
				}
				int maxIndex2 = page.MaxIndex;
				int num4 = page.RowCount - num;
				page.RowCount -= num4;
				return maxIndex2 + 1;
			}
			int rows = toRow - fromRow;
			if (shift)
			{
				UpdateRowIndex(page, num3, rows);
			}
			Array.Copy(page.Rows, num3, page.Rows, num, page.RowCount - num3);
			page.RowCount -= num3 - num;
			return toRow;
		}
		if (shift)
		{
			UpdateRowIndex(page, num3, toRow - fromRow);
		}
		if (toRow >= maxIndex)
		{
			return maxIndex;
		}
		return toRow;
	}

	private static void UpdateRowIndex(PageIndex page, int toPos, int rows)
	{
		for (int i = toPos; i < page.RowCount; i++)
		{
			page.Rows[i].Index -= (short)rows;
		}
	}

	private void DeleteColumns(int fromCol, int columns, bool shift)
	{
		int num = GetPosition(fromCol);
		if (num < 0)
		{
			num = ~num;
		}
		int num2 = num;
		for (int i = num; i <= ColumnCount; i++)
		{
			num2 = i;
			if (num2 == ColumnCount || _columnIndex[i].Index >= fromCol + columns)
			{
				break;
			}
		}
		if (ColumnCount <= num)
		{
			return;
		}
		if (_columnIndex[num].Index >= fromCol && _columnIndex[num].Index <= fromCol + columns)
		{
			if (num2 < ColumnCount)
			{
				Array.Copy(_columnIndex, num2, _columnIndex, num, ColumnCount - num2);
			}
			ColumnCount -= num2 - num;
		}
		if (shift)
		{
			for (int j = num; j < ColumnCount; j++)
			{
				_columnIndex[j].Index -= (short)columns;
			}
		}
	}

	private void UpdateIndexOffset(ColumnIndex column, int pagePos, int rowPos, int row, int rows)
	{
		if (pagePos >= column.PageCount)
		{
			return;
		}
		PageIndex pageIndex = column._pages[pagePos];
		if (rows > 1024)
		{
			short num = (short)(rows >> 10);
			int num2 = rows - 1024 * num;
			for (int i = pagePos + 1; i < column.PageCount; i++)
			{
				if (column._pages[i].Offset + num2 > 1024)
				{
					column._pages[i].Index += (short)(num + 1);
					column._pages[i].Offset += num2 - 1024;
				}
				else
				{
					column._pages[i].Index += num;
					column._pages[i].Offset += num2;
				}
			}
			int num3 = pageIndex.RowCount - rowPos;
			if (pageIndex.RowCount <= rowPos)
			{
				return;
			}
			if (column.PageCount - 1 == pagePos)
			{
				PageIndex pageIndex2 = CopyNew(pageIndex, rowPos, num3);
				pageIndex2.Index = (short)(row + rows >> 10);
				pageIndex2.Offset = row + rows - pageIndex2.Index * 1024 - pageIndex2.Rows[0].Index;
				if (pageIndex2.Offset > 1024)
				{
					pageIndex2.Index++;
					pageIndex2.Offset -= 1024;
				}
				AddPage(column, pagePos + 1, pageIndex2);
				pageIndex.RowCount = rowPos;
			}
			else if (column._pages[pagePos + 1].RowCount + num3 > 2048)
			{
				SplitPageInsert(column, pagePos, rowPos, rows, num3, num);
			}
			else
			{
				CopyMergePage(pageIndex, rowPos, rows, num3, column._pages[pagePos + 1]);
			}
			return;
		}
		for (int j = rowPos; j < pageIndex.RowCount; j++)
		{
			pageIndex.Rows[j].Index += (short)rows;
		}
		if (pageIndex.Offset + pageIndex.Rows[pageIndex.RowCount - 1].Index >= 2048)
		{
			AdjustIndex(column, pagePos);
			if (pageIndex.Offset + pageIndex.Rows[pageIndex.RowCount - 1].Index >= 2048)
			{
				pagePos = SplitPage(column, pagePos);
			}
		}
		for (int k = pagePos + 1; k < column.PageCount; k++)
		{
			if (column._pages[k].Offset + rows < 1024)
			{
				column._pages[k].Offset += rows;
				continue;
			}
			column._pages[k].Index++;
			column._pages[k].Offset = (column._pages[k].Offset + rows) % 1024;
		}
	}

	private void SplitPageInsert(ColumnIndex column, int pagePos, int rowPos, int rows, int size, int addPages)
	{
		_ = new IndexItem[GetSize(size)];
		PageIndex pageIndex = column._pages[pagePos];
		int num = -1;
		for (int i = rowPos; i < pageIndex.RowCount; i++)
		{
			if (pageIndex.IndexExpanded - (pageIndex.Rows[i].Index + rows) > 1024)
			{
				num = i;
				break;
			}
			pageIndex.Rows[i].Index += (short)rows;
		}
		int num2 = pageIndex.RowCount - num;
		pageIndex.RowCount = num;
		if (num2 > 0)
		{
			_ = pageIndex.IndexOffset;
			PageIndex pageIndex2 = CopyNew(pageIndex, num, num2);
			short num3 = (short)(pageIndex.Index + addPages);
			int num4 = pageIndex.IndexOffset + rows - num3 * 1024;
			if (num4 > 1024)
			{
				num3 += (short)(num4 / 1024);
				num4 %= 1024;
			}
			pageIndex2.Index = num3;
			pageIndex2.Offset = num4;
			AddPage(column, pagePos + 1, pageIndex2);
		}
	}

	private void CopyMergePage(PageIndex page, int rowPos, int rows, int size, PageIndex ToPage)
	{
		_ = page.IndexOffset;
		_ = ref page.Rows[rowPos];
		IndexItem[] array = new IndexItem[GetSize(ToPage.RowCount + size)];
		page.RowCount -= size;
		Array.Copy(page.Rows, rowPos, array, 0, size);
		for (int i = 0; i < size; i++)
		{
			array[i].Index += (short)(page.IndexOffset + rows - ToPage.IndexOffset);
		}
		Array.Copy(ToPage.Rows, 0, array, size, ToPage.RowCount);
		ToPage.Rows = array;
		ToPage.RowCount += size;
	}

	private void MergePage(ColumnIndex column, int pagePos)
	{
		PageIndex pageIndex = column._pages[pagePos];
		PageIndex pageIndex2 = column._pages[pagePos + 1];
		PageIndex pageIndex3 = new PageIndex(pageIndex, 0, pageIndex.RowCount + pageIndex2.RowCount);
		pageIndex3.RowCount = pageIndex.RowCount + pageIndex2.RowCount;
		Array.Copy(pageIndex.Rows, 0, pageIndex3.Rows, 0, pageIndex.RowCount);
		Array.Copy(pageIndex2.Rows, 0, pageIndex3.Rows, pageIndex.RowCount, pageIndex2.RowCount);
		for (int i = pageIndex.RowCount; i < pageIndex3.RowCount; i++)
		{
			pageIndex3.Rows[i].Index += (short)(pageIndex2.IndexOffset - pageIndex.IndexOffset);
		}
		column._pages[pagePos] = pageIndex3;
		column.PageCount--;
		if (column.PageCount > pagePos + 1)
		{
			Array.Copy(column._pages, pagePos + 2, column._pages, pagePos + 1, column.PageCount - (pagePos + 1));
			for (int j = pagePos + 1; j < column.PageCount; j++)
			{
				column._pages[j].Index--;
				column._pages[j].Offset += 1024;
			}
		}
	}

	private PageIndex CopyNew(PageIndex pageFrom, int rowPos, int size)
	{
		IndexItem[] array = new IndexItem[GetSize(size)];
		Array.Copy(pageFrom.Rows, rowPos, array, 0, size);
		return new PageIndex(array, size);
	}

	internal static int GetSize(int size)
	{
		int num;
		for (num = 256; num < size; num <<= 1)
		{
		}
		return num;
	}

	private void AddCell(ColumnIndex columnIndex, int pagePos, int pos, short ix, T value)
	{
		PageIndex pageIndex = columnIndex._pages[pagePos];
		if (pageIndex.RowCount == pageIndex.Rows.Length)
		{
			if (pageIndex.RowCount == 2048)
			{
				pagePos = SplitPage(columnIndex, pagePos);
				if (columnIndex._pages[pagePos - 1].RowCount > pos)
				{
					pagePos--;
				}
				else
				{
					pos -= columnIndex._pages[pagePos - 1].RowCount;
				}
				pageIndex = columnIndex._pages[pagePos];
			}
			else
			{
				IndexItem[] array = new IndexItem[pageIndex.Rows.Length << 1];
				Array.Copy(pageIndex.Rows, 0, array, 0, pageIndex.RowCount);
				pageIndex.Rows = array;
			}
		}
		if (pos < pageIndex.RowCount)
		{
			Array.Copy(pageIndex.Rows, pos, pageIndex.Rows, pos + 1, pageIndex.RowCount - pos);
		}
		pageIndex.Rows[pos] = new IndexItem
		{
			Index = ix,
			IndexPointer = _values.Count
		};
		_values.Add(value);
		pageIndex.RowCount++;
	}

	private int SplitPage(ColumnIndex columnIndex, int pagePos)
	{
		PageIndex pageIndex = columnIndex._pages[pagePos];
		if (pageIndex.Offset != 0)
		{
			int offset = pageIndex.Offset;
			pageIndex.Offset = 0;
			for (int i = 0; i < pageIndex.RowCount; i++)
			{
				pageIndex.Rows[i].Index -= (short)offset;
			}
		}
		int num = 0;
		for (int j = 0; j < pageIndex.RowCount; j++)
		{
			if (pageIndex.Rows[j].Index > 1024)
			{
				num = j;
				break;
			}
		}
		PageIndex pageIndex2 = new PageIndex(pageIndex, 0, num);
		PageIndex pageIndex3 = new PageIndex(pageIndex, num, pageIndex.RowCount - num, (short)(pageIndex.Index + 1), pageIndex.Offset);
		for (int k = 0; k < pageIndex3.RowCount; k++)
		{
			pageIndex3.Rows[k].Index = (short)(pageIndex3.Rows[k].Index - 1024);
		}
		columnIndex._pages[pagePos] = pageIndex2;
		if (columnIndex.PageCount + 1 > columnIndex._pages.Length)
		{
			PageIndex[] array = new PageIndex[columnIndex._pages.Length << 1];
			Array.Copy(columnIndex._pages, 0, array, 0, columnIndex.PageCount);
			columnIndex._pages = array;
		}
		Array.Copy(columnIndex._pages, pagePos + 1, columnIndex._pages, pagePos + 2, columnIndex.PageCount - pagePos - 1);
		columnIndex._pages[pagePos + 1] = pageIndex3;
		pageIndex = pageIndex3;
		columnIndex.PageCount++;
		return pagePos + 1;
	}

	private PageIndex AdjustIndex(ColumnIndex columnIndex, int pagePos)
	{
		PageIndex pageIndex = columnIndex._pages[pagePos];
		if (pageIndex.Offset + pageIndex.Rows[0].Index >= 1024 || pageIndex.Offset >= 1024 || pageIndex.Rows[0].Index >= 1024)
		{
			pageIndex.Index++;
			pageIndex.Offset -= 1024;
		}
		else if (pageIndex.Offset + pageIndex.Rows[0].Index <= -1024 || pageIndex.Offset <= -1024 || pageIndex.Rows[0].Index <= -1024)
		{
			pageIndex.Index--;
			pageIndex.Offset += 1024;
		}
		return pageIndex;
	}

	private void AddPageRowOffset(PageIndex page, short offset)
	{
		for (int i = 0; i < page.RowCount; i++)
		{
			page.Rows[i].Index += offset;
		}
	}

	private void AddPage(ColumnIndex column, int pos, short index)
	{
		AddPage(column, pos);
		column._pages[pos] = new PageIndex
		{
			Index = index
		};
		if (pos > 0)
		{
			PageIndex pageIndex = column._pages[pos - 1];
			if (pageIndex.RowCount > 0 && pageIndex.Rows[pageIndex.RowCount - 1].Index > 1024)
			{
				column._pages[pos].Offset = pageIndex.Rows[pageIndex.RowCount - 1].Index - 1024;
			}
		}
	}

	private void AddPage(ColumnIndex column, int pos, PageIndex page)
	{
		AddPage(column, pos);
		column._pages[pos] = page;
	}

	private void AddPage(ColumnIndex column, int pos)
	{
		if (column.PageCount == column._pages.Length)
		{
			PageIndex[] array = new PageIndex[column._pages.Length * 2];
			Array.Copy(column._pages, 0, array, 0, column.PageCount);
			column._pages = array;
		}
		if (pos < column.PageCount)
		{
			Array.Copy(column._pages, pos, column._pages, pos + 1, column.PageCount - pos);
		}
		column.PageCount++;
	}

	private void AddColumn(int pos, int Column)
	{
		if (ColumnCount == _columnIndex.Length)
		{
			ColumnIndex[] array = new ColumnIndex[_columnIndex.Length * 2];
			Array.Copy(_columnIndex, 0, array, 0, ColumnCount);
			_columnIndex = array;
		}
		if (pos < ColumnCount)
		{
			Array.Copy(_columnIndex, pos, _columnIndex, pos + 1, ColumnCount - pos);
		}
		_columnIndex[pos] = new ColumnIndex
		{
			Index = (short)Column
		};
		ColumnCount++;
	}

	public void Dispose()
	{
		if (_values != null)
		{
			_values.Clear();
		}
		for (int i = 0; i < ColumnCount; i++)
		{
			if (_columnIndex[i] != null)
			{
				((IDisposable)_columnIndex[i]).Dispose();
			}
		}
		_values = null;
		_columnIndex = null;
	}

	public bool MoveNext()
	{
		return GetNextCell(ref _row, ref _colPos, 0, 1048576, 16384);
	}

	internal bool NextCell(ref int row, ref int col)
	{
		return NextCell(ref row, ref col, 0, 0, 1048576, 16384);
	}

	internal bool NextCell(ref int row, ref int col, int minRow, int minColPos, int maxRow, int maxColPos)
	{
		if (minColPos >= ColumnCount)
		{
			return false;
		}
		if (maxColPos >= ColumnCount)
		{
			maxColPos = ColumnCount - 1;
		}
		int colPos = GetPosition(col);
		if (colPos >= 0)
		{
			if (colPos > maxColPos)
			{
				if (col <= minColPos)
				{
					return false;
				}
				col = minColPos;
				return NextCell(ref row, ref col);
			}
			bool nextCell = GetNextCell(ref row, ref colPos, minColPos, maxRow, maxColPos);
			col = _columnIndex[colPos].Index;
			return nextCell;
		}
		colPos = ~colPos;
		if (colPos >= ColumnCount)
		{
			colPos = ColumnCount - 1;
		}
		if (col > _columnIndex[colPos].Index)
		{
			if (col <= minColPos)
			{
				return false;
			}
			col = minColPos;
			return NextCell(ref row, ref col, minRow, minColPos, maxRow, maxColPos);
		}
		bool nextCell2 = GetNextCell(ref row, ref colPos, minColPos, maxRow, maxColPos);
		col = _columnIndex[colPos].Index;
		return nextCell2;
	}

	internal bool GetNextCell(ref int row, ref int colPos, int startColPos, int endRow, int endColPos)
	{
		if (ColumnCount == 0)
		{
			return false;
		}
		if (++colPos < ColumnCount && colPos <= endColPos)
		{
			int nextRow = _columnIndex[colPos].GetNextRow(row);
			if (nextRow == row)
			{
				return true;
			}
			int num;
			int num2;
			if (nextRow > row)
			{
				num = nextRow;
				num2 = colPos;
			}
			else
			{
				num = int.MaxValue;
				num2 = 0;
			}
			int i;
			for (i = colPos + 1; i < ColumnCount && i <= endColPos; i++)
			{
				nextRow = _columnIndex[i].GetNextRow(row);
				if (nextRow == row)
				{
					colPos = i;
					return true;
				}
				if (nextRow > row && nextRow < num)
				{
					num = nextRow;
					num2 = i;
				}
			}
			i = startColPos;
			if (row < endRow)
			{
				row++;
				for (; i < colPos; i++)
				{
					nextRow = _columnIndex[i].GetNextRow(row);
					if (nextRow == row)
					{
						colPos = i;
						return true;
					}
					if (nextRow > row && (nextRow < num || (nextRow == num && i < num2)) && nextRow <= endRow)
					{
						num = nextRow;
						num2 = i;
					}
				}
			}
			if (num == int.MaxValue || num > endRow)
			{
				return false;
			}
			row = num;
			colPos = num2;
			return true;
		}
		if (colPos <= startColPos || row >= endRow)
		{
			return false;
		}
		colPos = startColPos - 1;
		row++;
		return GetNextCell(ref row, ref colPos, startColPos, endRow, endColPos);
	}

	internal bool GetNextCell(ref int row, ref int colPos, int startColPos, int endRow, int endColPos, ref int[] pagePos, ref int[] cellPos)
	{
		if (colPos == endColPos)
		{
			colPos = startColPos;
			row++;
		}
		else
		{
			colPos++;
		}
		if (pagePos[colPos] < 0)
		{
			if (pagePos[colPos] == -1)
			{
				pagePos[colPos] = _columnIndex[colPos].GetPosition(row);
			}
		}
		else if (_columnIndex[colPos]._pages[pagePos[colPos]].RowCount <= row)
		{
			if (_columnIndex[colPos].PageCount > pagePos[colPos])
			{
				pagePos[colPos]++;
			}
			else
			{
				pagePos[colPos] = -2;
			}
		}
		int num = _columnIndex[colPos]._pages[pagePos[colPos]].IndexOffset + _columnIndex[colPos]._pages[pagePos[colPos]].Rows[cellPos[colPos]].Index;
		if (num == row)
		{
			row = num;
		}
		return true;
	}

	internal bool PrevCell(ref int row, ref int col)
	{
		return PrevCell(ref row, ref col, 0, 0, 1048576, 16384);
	}

	internal bool PrevCell(ref int row, ref int col, int minRow, int minColPos, int maxRow, int maxColPos)
	{
		if (minColPos >= ColumnCount)
		{
			return false;
		}
		if (maxColPos >= ColumnCount)
		{
			maxColPos = ColumnCount - 1;
		}
		int colPos = GetPosition(col);
		if (colPos >= 0)
		{
			if (colPos == 0)
			{
				if (col >= maxColPos)
				{
					return false;
				}
				if (row == minRow)
				{
					return false;
				}
				row--;
				col = maxColPos;
				return PrevCell(ref row, ref col, minRow, minColPos, maxRow, maxColPos);
			}
			bool prevCell = GetPrevCell(ref row, ref colPos, minRow, minColPos, maxColPos);
			if (prevCell)
			{
				col = _columnIndex[colPos].Index;
			}
			return prevCell;
		}
		colPos = ~colPos;
		if (colPos == 0)
		{
			if (col >= maxColPos || row <= 0)
			{
				return false;
			}
			col = maxColPos;
			row--;
			return PrevCell(ref row, ref col, minRow, minColPos, maxRow, maxColPos);
		}
		bool prevCell2 = GetPrevCell(ref row, ref colPos, minRow, minColPos, maxColPos);
		if (prevCell2)
		{
			col = _columnIndex[colPos].Index;
		}
		return prevCell2;
	}

	internal bool GetPrevCell(ref int row, ref int colPos, int startRow, int startColPos, int endColPos)
	{
		if (ColumnCount == 0)
		{
			return false;
		}
		if (--colPos >= startColPos)
		{
			int nextRow = _columnIndex[colPos].GetNextRow(row);
			if (nextRow == row)
			{
				return true;
			}
			int num;
			int num2;
			if (nextRow > row && nextRow >= startRow)
			{
				num = nextRow;
				num2 = colPos;
			}
			else
			{
				num = int.MaxValue;
				num2 = 0;
			}
			int num3 = colPos - 1;
			if (num3 >= startColPos)
			{
				while (num3 >= startColPos)
				{
					nextRow = _columnIndex[num3].GetNextRow(row);
					if (nextRow == row)
					{
						colPos = num3;
						return true;
					}
					if (nextRow > row && nextRow < num && nextRow >= startRow)
					{
						num = nextRow;
						num2 = num3;
					}
					num3--;
				}
			}
			if (row > startRow)
			{
				num3 = endColPos;
				row--;
				while (num3 > colPos)
				{
					nextRow = _columnIndex[num3].GetNextRow(row);
					if (nextRow == row)
					{
						colPos = num3;
						return true;
					}
					if (nextRow > row && nextRow < num && nextRow >= startRow)
					{
						num = nextRow;
						num2 = num3;
					}
					num3--;
				}
			}
			if (num == int.MaxValue || startRow < num)
			{
				return false;
			}
			row = num;
			colPos = num2;
			return true;
		}
		colPos = ColumnCount;
		row--;
		if (row < startRow)
		{
			Reset();
			return false;
		}
		return GetPrevCell(ref colPos, ref row, startRow, startColPos, endColPos);
	}

	public void Reset()
	{
		_colPos = -1;
		_row = 0;
	}
}
