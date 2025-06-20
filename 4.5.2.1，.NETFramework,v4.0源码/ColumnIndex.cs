using System;

internal class ColumnIndex : IndexBase, IDisposable
{
	internal IndexBase _searchIx = new IndexBase();

	internal PageIndex[] _pages;

	internal int PageCount;

	public ColumnIndex()
	{
		_pages = new PageIndex[32];
		PageCount = 0;
	}

	~ColumnIndex()
	{
		_pages = null;
	}

	internal int GetPosition(int Row)
	{
		short num = (short)(Row >> 10);
		int res;
		if (num >= 0 && num < PageCount && _pages[num].Index == num)
		{
			res = num;
		}
		else
		{
			_searchIx.Index = num;
			res = Array.BinarySearch(_pages, 0, PageCount, _searchIx);
		}
		if (res >= 0)
		{
			GetPage(Row, ref res);
			return res;
		}
		int res2 = ~res;
		if (GetPage(Row, ref res2))
		{
			return res2;
		}
		return res;
	}

	private bool GetPage(int Row, ref int res)
	{
		if (res < PageCount && _pages[res].MinIndex <= Row && _pages[res].MaxIndex >= Row)
		{
			return true;
		}
		if (res + 1 < PageCount && _pages[res + 1].MinIndex <= Row)
		{
			do
			{
				res++;
			}
			while (res + 1 < PageCount && _pages[res + 1].MinIndex <= Row);
			return true;
		}
		if (res - 1 >= 0 && _pages[res - 1].MaxIndex >= Row)
		{
			do
			{
				res--;
			}
			while (res - 1 > 0 && _pages[res - 1].MaxIndex >= Row);
			return true;
		}
		return false;
	}

	internal int GetNextRow(int row)
	{
		int position = GetPosition(row);
		if (position < 0)
		{
			position = ~position;
			if (position >= PageCount)
			{
				return -1;
			}
			if (_pages[position].IndexOffset + _pages[position].Rows[0].Index < row)
			{
				if (position + 1 >= PageCount)
				{
					return -1;
				}
				return _pages[position + 1].IndexOffset + _pages[position].Rows[0].Index;
			}
			return _pages[position].IndexOffset + _pages[position].Rows[0].Index;
		}
		if (position < PageCount)
		{
			int nextRow = _pages[position].GetNextRow(row);
			if (nextRow >= 0)
			{
				return _pages[position].IndexOffset + _pages[position].Rows[nextRow].Index;
			}
			if (++position < PageCount)
			{
				return _pages[position].IndexOffset + _pages[position].Rows[0].Index;
			}
			return -1;
		}
		return -1;
	}

	internal int FindNext(int Page)
	{
		int position = GetPosition(Page);
		if (position < 0)
		{
			return ~position;
		}
		return position;
	}

	public void Dispose()
	{
		for (int i = 0; i < PageCount; i++)
		{
			((IDisposable)_pages[i]).Dispose();
		}
		_pages = null;
	}
}
