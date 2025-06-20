using System;
using System.Drawing;

namespace OfficeOpenXml.Style;

public sealed class ExcelFont : StyleBase
{
	public string Name
	{
		get
		{
			return _styles.Fonts[base.Index].Name;
		}
		set
		{
			_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Font, eStyleProperty.Name, value, _positionID, _address));
		}
	}

	public float Size
	{
		get
		{
			return _styles.Fonts[base.Index].Size;
		}
		set
		{
			_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Font, eStyleProperty.Size, value, _positionID, _address));
		}
	}

	public int Family
	{
		get
		{
			return _styles.Fonts[base.Index].Family;
		}
		set
		{
			_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Font, eStyleProperty.Family, value, _positionID, _address));
		}
	}

	public ExcelColor Color => new ExcelColor(_styles, _ChangedEvent, _positionID, _address, eStyleClass.Font, this);

	public string Scheme
	{
		get
		{
			return _styles.Fonts[base.Index].Scheme;
		}
		set
		{
			_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Font, eStyleProperty.Scheme, value, _positionID, _address));
		}
	}

	public bool Bold
	{
		get
		{
			return _styles.Fonts[base.Index].Bold;
		}
		set
		{
			_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Font, eStyleProperty.Bold, value, _positionID, _address));
		}
	}

	public bool Italic
	{
		get
		{
			return _styles.Fonts[base.Index].Italic;
		}
		set
		{
			_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Font, eStyleProperty.Italic, value, _positionID, _address));
		}
	}

	public bool Strike
	{
		get
		{
			return _styles.Fonts[base.Index].Strike;
		}
		set
		{
			_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Font, eStyleProperty.Strike, value, _positionID, _address));
		}
	}

	public bool UnderLine
	{
		get
		{
			return _styles.Fonts[base.Index].UnderLine;
		}
		set
		{
			if (value)
			{
				UnderLineType = ExcelUnderLineType.Single;
			}
			else
			{
				UnderLineType = ExcelUnderLineType.None;
			}
		}
	}

	public ExcelUnderLineType UnderLineType
	{
		get
		{
			return _styles.Fonts[base.Index].UnderLineType;
		}
		set
		{
			_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Font, eStyleProperty.UnderlineType, value, _positionID, _address));
		}
	}

	public ExcelVerticalAlignmentFont VerticalAlign
	{
		get
		{
			if (_styles.Fonts[base.Index].VerticalAlign == "")
			{
				return ExcelVerticalAlignmentFont.None;
			}
			return (ExcelVerticalAlignmentFont)Enum.Parse(typeof(ExcelVerticalAlignmentFont), _styles.Fonts[base.Index].VerticalAlign, ignoreCase: true);
		}
		set
		{
			_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Font, eStyleProperty.VerticalAlign, value, _positionID, _address));
		}
	}

	internal override string Id => Name + Size.ToString() + Family.ToString() + Scheme.ToString() + Bold.ToString()[0].ToString() + Italic.ToString()[0].ToString() + Strike.ToString()[0].ToString() + UnderLine.ToString()[0].ToString() + VerticalAlign;

	internal ExcelFont(ExcelStyles styles, XmlHelper.ChangedEventHandler ChangedEvent, int PositionID, string address, int index)
		: base(styles, ChangedEvent, PositionID, address)
	{
		base.Index = index;
	}

	public void SetFromFont(Font Font)
	{
		Name = Font.Name;
		Size = (int)Font.Size;
		Strike = Font.Strikeout;
		Bold = Font.Bold;
		UnderLine = Font.Underline;
		Italic = Font.Italic;
	}
}
