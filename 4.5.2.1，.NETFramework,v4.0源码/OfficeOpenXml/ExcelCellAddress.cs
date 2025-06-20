using System;

namespace OfficeOpenXml;

public class ExcelCellAddress
{
	private int _row;

	private int _column;

	private string _address;

	public int Row
	{
		get
		{
			return _row;
		}
		private set
		{
			if (value <= 0)
			{
				throw new ArgumentOutOfRangeException("value", "Row cannot be less than 1.");
			}
			_row = value;
			if (_column > 0)
			{
				_address = ExcelCellBase.GetAddress(_row, _column);
			}
			else
			{
				_address = "#REF!";
			}
		}
	}

	public int Column
	{
		get
		{
			return _column;
		}
		private set
		{
			if (value <= 0)
			{
				throw new ArgumentOutOfRangeException("value", "Column cannot be less than 1.");
			}
			_column = value;
			if (_row > 0)
			{
				_address = ExcelCellBase.GetAddress(_row, _column);
			}
			else
			{
				_address = "#REF!";
			}
		}
	}

	public string Address
	{
		get
		{
			return _address;
		}
		internal set
		{
			_address = value;
			ExcelCellBase.GetRowColFromAddress(_address, out _row, out _column);
		}
	}

	public bool IsRef => _row <= 0;

	public ExcelCellAddress()
		: this(1, 1)
	{
	}

	public ExcelCellAddress(int row, int column)
	{
		Row = row;
		Column = column;
	}

	public ExcelCellAddress(string address)
	{
		Address = address;
	}

	public static string GetColumnLetter(int column)
	{
		if (column > 16384 || column < 1)
		{
			throw new InvalidOperationException("Invalid 1-based column index: " + column + ". Valid range is 1 to " + 16384);
		}
		return ExcelCellBase.GetColumnLetter(column);
	}
}
