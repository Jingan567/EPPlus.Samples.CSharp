using System;

internal class PageIndex : IndexBase, IDisposable
{
	internal IndexItem _searchIx;

	internal int Offset;

	internal int RowCount;

	internal int IndexOffset => IndexExpanded + Offset;

	internal int IndexExpanded => Index << 10;

	internal IndexItem[] Rows { get; set; }

	public int MinIndex
	{
		get
		{
			if (Rows.Length != 0)
			{
				return IndexOffset + Rows[0].Index;
			}
			return -1;
		}
	}

	public int MaxIndex
	{
		get
		{
			if (RowCount > 0)
			{
				return IndexOffset + Rows[RowCount - 1].Index;
			}
			return -1;
		}
	}

	public PageIndex()
	{
		Rows = new IndexItem[1024];
		RowCount = 0;
	}

	public PageIndex(IndexItem[] rows, int count)
	{
		Rows = rows;
		RowCount = count;
	}

	public PageIndex(PageIndex pageItem, int start, int size)
		: this(pageItem, start, size, pageItem.Index, pageItem.Offset)
	{
	}

	public PageIndex(PageIndex pageItem, int start, int size, short index, int offset)
	{
		Rows = new IndexItem[CellStore<int>.GetSize(size)];
		Array.Copy(pageItem.Rows, start, Rows, 0, size);
		RowCount = size;
		Index = index;
		Offset = offset;
	}

	~PageIndex()
	{
		Rows = null;
	}

	internal int GetPosition(int offset)
	{
		_searchIx.Index = (short)offset;
		return Array.BinarySearch(Rows, 0, RowCount, _searchIx);
	}

	internal int GetNextRow(int row)
	{
		int offset = row - IndexOffset;
		int position = GetPosition(offset);
		if (position < 0)
		{
			position = ~position;
			if (position < RowCount)
			{
				return position;
			}
			return -1;
		}
		return position;
	}

	public int GetIndex(int pos)
	{
		return IndexOffset + Rows[pos].Index;
	}

	public void Dispose()
	{
		Rows = null;
	}
}
