using System;
using OfficeOpenXml.Style.XmlAccess;

namespace OfficeOpenXml.Style;

public sealed class ExcelStyle : StyleBase
{
	private const string xfIdPath = "@xfId";

	public ExcelNumberFormat Numberformat { get; set; }

	public ExcelFont Font { get; set; }

	public ExcelFill Fill { get; set; }

	public Border Border { get; set; }

	public ExcelHorizontalAlignment HorizontalAlignment
	{
		get
		{
			return _styles.CellXfs[base.Index].HorizontalAlignment;
		}
		set
		{
			_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.HorizontalAlign, value, _positionID, _address));
		}
	}

	public ExcelVerticalAlignment VerticalAlignment
	{
		get
		{
			return _styles.CellXfs[base.Index].VerticalAlignment;
		}
		set
		{
			_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.VerticalAlign, value, _positionID, _address));
		}
	}

	public bool WrapText
	{
		get
		{
			return _styles.CellXfs[base.Index].WrapText;
		}
		set
		{
			_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.WrapText, value, _positionID, _address));
		}
	}

	public ExcelReadingOrder ReadingOrder
	{
		get
		{
			return _styles.CellXfs[base.Index].ReadingOrder;
		}
		set
		{
			_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.ReadingOrder, value, _positionID, _address));
		}
	}

	public bool ShrinkToFit
	{
		get
		{
			return _styles.CellXfs[base.Index].ShrinkToFit;
		}
		set
		{
			_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.ShrinkToFit, value, _positionID, _address));
		}
	}

	public int Indent
	{
		get
		{
			return _styles.CellXfs[base.Index].Indent;
		}
		set
		{
			if (value < 0 || value > 250)
			{
				throw new ArgumentOutOfRangeException("Indent must be between 0 and 250");
			}
			_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.Indent, value, _positionID, _address));
		}
	}

	public int TextRotation
	{
		get
		{
			return _styles.CellXfs[base.Index].TextRotation;
		}
		set
		{
			if (value < 0 || value > 180)
			{
				throw new ArgumentOutOfRangeException("TextRotation out of range.");
			}
			_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.TextRotation, value, _positionID, _address));
		}
	}

	public bool Locked
	{
		get
		{
			return _styles.CellXfs[base.Index].Locked;
		}
		set
		{
			_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.Locked, value, _positionID, _address));
		}
	}

	public bool Hidden
	{
		get
		{
			return _styles.CellXfs[base.Index].Hidden;
		}
		set
		{
			_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.Hidden, value, _positionID, _address));
		}
	}

	public bool QuotePrefix
	{
		get
		{
			return _styles.CellXfs[base.Index].QuotePrefix;
		}
		set
		{
			_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.QuotePrefix, value, _positionID, _address));
		}
	}

	public int XfId
	{
		get
		{
			return _styles.CellXfs[base.Index].XfId;
		}
		set
		{
			_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.XfId, value, _positionID, _address));
		}
	}

	internal int PositionID { get; set; }

	internal ExcelStyles Styles { get; set; }

	internal override string Id => string.Concat(Numberformat.Id, "|", Font.Id, "|", Fill.Id, "|", Border.Id, "|", VerticalAlignment, "|", HorizontalAlignment, "|", WrapText.ToString(), "|", ReadingOrder.ToString(), "|", XfId.ToString(), "|", QuotePrefix.ToString());

	internal ExcelStyle(ExcelStyles styles, XmlHelper.ChangedEventHandler ChangedEvent, int positionID, string Address, int xfsId)
		: base(styles, ChangedEvent, positionID, Address)
	{
		base.Index = xfsId;
		ExcelXfs excelXfs = ((positionID <= -1) ? _styles.CellStyleXfs[xfsId] : _styles.CellXfs[xfsId]);
		Styles = styles;
		PositionID = positionID;
		Numberformat = new ExcelNumberFormat(styles, ChangedEvent, PositionID, Address, excelXfs.NumberFormatId);
		Font = new ExcelFont(styles, ChangedEvent, PositionID, Address, excelXfs.FontId);
		Fill = new ExcelFill(styles, ChangedEvent, PositionID, Address, excelXfs.FillId);
		Border = new Border(styles, ChangedEvent, PositionID, Address, excelXfs.BorderId);
	}
}
