using System.Drawing;

namespace OfficeOpenXml.Style;

public sealed class Border : StyleBase
{
	public ExcelBorderItem Left => new ExcelBorderItem(_styles, _ChangedEvent, _positionID, _address, eStyleClass.BorderLeft, this);

	public ExcelBorderItem Right => new ExcelBorderItem(_styles, _ChangedEvent, _positionID, _address, eStyleClass.BorderRight, this);

	public ExcelBorderItem Top => new ExcelBorderItem(_styles, _ChangedEvent, _positionID, _address, eStyleClass.BorderTop, this);

	public ExcelBorderItem Bottom => new ExcelBorderItem(_styles, _ChangedEvent, _positionID, _address, eStyleClass.BorderBottom, this);

	public ExcelBorderItem Diagonal => new ExcelBorderItem(_styles, _ChangedEvent, _positionID, _address, eStyleClass.BorderDiagonal, this);

	public bool DiagonalUp
	{
		get
		{
			if (base.Index >= 0)
			{
				return _styles.Borders[base.Index].DiagonalUp;
			}
			return false;
		}
		set
		{
			_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Border, eStyleProperty.BorderDiagonalUp, value, _positionID, _address));
		}
	}

	public bool DiagonalDown
	{
		get
		{
			if (base.Index >= 0)
			{
				return _styles.Borders[base.Index].DiagonalDown;
			}
			return false;
		}
		set
		{
			_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Border, eStyleProperty.BorderDiagonalDown, value, _positionID, _address));
		}
	}

	internal override string Id => Top.Id + Bottom.Id + Left.Id + Right.Id + Diagonal.Id + DiagonalUp + DiagonalDown;

	internal Border(ExcelStyles styles, XmlHelper.ChangedEventHandler ChangedEvent, int PositionID, string address, int index)
		: base(styles, ChangedEvent, PositionID, address)
	{
		base.Index = index;
	}

	public void BorderAround(ExcelBorderStyle Style)
	{
		ExcelAddress addr = new ExcelAddress(_address);
		SetBorderAroundStyle(Style, addr);
	}

	public void BorderAround(ExcelBorderStyle Style, Color Color)
	{
		ExcelAddress excelAddress = new ExcelAddress(_address);
		SetBorderAroundStyle(Style, excelAddress);
		_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.BorderTop, eStyleProperty.Color, Color.ToArgb().ToString("X"), _positionID, new ExcelAddress(excelAddress._fromRow, excelAddress._fromCol, excelAddress._fromRow, excelAddress._toCol).Address));
		_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.BorderBottom, eStyleProperty.Color, Color.ToArgb().ToString("X"), _positionID, new ExcelAddress(excelAddress._toRow, excelAddress._fromCol, excelAddress._toRow, excelAddress._toCol).Address));
		_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.BorderLeft, eStyleProperty.Color, Color.ToArgb().ToString("X"), _positionID, new ExcelAddress(excelAddress._fromRow, excelAddress._fromCol, excelAddress._toRow, excelAddress._fromCol).Address));
		_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.BorderRight, eStyleProperty.Color, Color.ToArgb().ToString("X"), _positionID, new ExcelAddress(excelAddress._fromRow, excelAddress._toCol, excelAddress._toRow, excelAddress._toCol).Address));
	}

	private void SetBorderAroundStyle(ExcelBorderStyle Style, ExcelAddress addr)
	{
		_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.BorderTop, eStyleProperty.Style, Style, _positionID, new ExcelAddress(addr._fromRow, addr._fromCol, addr._fromRow, addr._toCol).Address));
		_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.BorderBottom, eStyleProperty.Style, Style, _positionID, new ExcelAddress(addr._toRow, addr._fromCol, addr._toRow, addr._toCol).Address));
		_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.BorderLeft, eStyleProperty.Style, Style, _positionID, new ExcelAddress(addr._fromRow, addr._fromCol, addr._toRow, addr._fromCol).Address));
		_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.BorderRight, eStyleProperty.Style, Style, _positionID, new ExcelAddress(addr._fromRow, addr._toCol, addr._toRow, addr._toCol).Address));
	}
}
