using System;

namespace OfficeOpenXml;

public class ExcelRange : ExcelRangeBase
{
	public ExcelRange this[string Address]
	{
		get
		{
			if (_worksheet.Names.ContainsKey(Address))
			{
				if (_worksheet.Names[Address].IsName)
				{
					return null;
				}
				base.Address = _worksheet.Names[Address].Address;
			}
			else
			{
				base.Address = Address;
			}
			_rtc = null;
			return this;
		}
	}

	public ExcelRange this[int Row, int Col]
	{
		get
		{
			ValidateRowCol(Row, Col);
			_fromCol = Col;
			_fromRow = Row;
			_toCol = Col;
			_toRow = Row;
			_rtc = null;
			_start = null;
			_end = null;
			_addresses = null;
			_address = ExcelCellBase.GetAddress(_fromRow, _fromCol);
			ChangeAddress();
			return this;
		}
	}

	public ExcelRange this[int FromRow, int FromCol, int ToRow, int ToCol]
	{
		get
		{
			ValidateRowCol(FromRow, FromCol);
			ValidateRowCol(ToRow, ToCol);
			_fromCol = FromCol;
			_fromRow = FromRow;
			_toCol = ToCol;
			_toRow = ToRow;
			_rtc = null;
			_start = null;
			_end = null;
			_addresses = null;
			_address = ExcelCellBase.GetAddress(_fromRow, _fromCol, _toRow, _toCol);
			ChangeAddress();
			return this;
		}
	}

	internal ExcelRange(ExcelWorksheet sheet)
		: base(sheet)
	{
	}

	internal ExcelRange(ExcelWorksheet sheet, string address)
		: base(sheet, address)
	{
	}

	internal ExcelRange(ExcelWorksheet sheet, int fromRow, int fromCol, int toRow, int toCol)
		: base(sheet)
	{
		_fromRow = fromRow;
		_fromCol = fromCol;
		_toRow = toRow;
		_toCol = toCol;
	}

	private ExcelRange GetTableAddess(ExcelWorksheet _worksheet, string address)
	{
		int num = address.IndexOf('[');
		if (num == 0)
		{
			int num2 = address.IndexOf(']', num + 1);
			if (num >= 0 && num2 >= 0)
			{
				address.Substring(num + 1, num2 - 1);
			}
		}
		return null;
	}

	private static void ValidateRowCol(int Row, int Col)
	{
		if (Row < 1 || Row > 1048576)
		{
			throw new ArgumentException("Row out of range");
		}
		if (Col < 1 || Col > 16384)
		{
			throw new ArgumentException("Column out of range");
		}
	}
}
