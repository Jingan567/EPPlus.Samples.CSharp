using System;
using OfficeOpenXml.Style.XmlAccess;

namespace OfficeOpenXml.Style;

public sealed class ExcelBorderItem : StyleBase
{
	private eStyleClass _cls;

	private StyleBase _parent;

	private ExcelColor _color;

	public ExcelBorderStyle Style
	{
		get
		{
			return GetSource().Style;
		}
		set
		{
			_ChangedEvent(this, new StyleChangeEventArgs(_cls, eStyleProperty.Style, value, _positionID, _address));
		}
	}

	public ExcelColor Color
	{
		get
		{
			if (_color == null)
			{
				_color = new ExcelColor(_styles, _ChangedEvent, _positionID, _address, _cls, _parent);
			}
			return _color;
		}
	}

	internal override string Id => string.Concat(Style, Color.Id);

	internal ExcelBorderItem(ExcelStyles styles, XmlHelper.ChangedEventHandler ChangedEvent, int worksheetID, string address, eStyleClass cls, StyleBase parent)
		: base(styles, ChangedEvent, worksheetID, address)
	{
		_cls = cls;
		_parent = parent;
	}

	internal override void SetIndex(int index)
	{
		_parent.Index = index;
	}

	private ExcelBorderItemXml GetSource()
	{
		int positionID = ((_parent.Index >= 0) ? _parent.Index : 0);
		return _cls switch
		{
			eStyleClass.BorderTop => _styles.Borders[positionID].Top, 
			eStyleClass.BorderBottom => _styles.Borders[positionID].Bottom, 
			eStyleClass.BorderLeft => _styles.Borders[positionID].Left, 
			eStyleClass.BorderRight => _styles.Borders[positionID].Right, 
			eStyleClass.BorderDiagonal => _styles.Borders[positionID].Diagonal, 
			_ => throw new Exception("Invalid class for Borderitem"), 
		};
	}
}
