using System;
using System.Drawing;
using OfficeOpenXml.Style.XmlAccess;

namespace OfficeOpenXml.Style;

public sealed class ExcelColor : StyleBase, IColor
{
	private eStyleClass _cls;

	private StyleBase _parent;

	public string Theme => GetSource().Theme;

	public decimal Tint
	{
		get
		{
			return GetSource().Tint;
		}
		set
		{
			if (value > 1m || value < -1m)
			{
				throw new ArgumentOutOfRangeException("Value must be between -1 and 1");
			}
			_ChangedEvent(this, new StyleChangeEventArgs(_cls, eStyleProperty.Tint, value, _positionID, _address));
		}
	}

	public string Rgb
	{
		get
		{
			return GetSource().Rgb;
		}
		internal set
		{
			_ChangedEvent(this, new StyleChangeEventArgs(_cls, eStyleProperty.Color, value, _positionID, _address));
		}
	}

	public int Indexed
	{
		get
		{
			return GetSource().Indexed;
		}
		set
		{
			_ChangedEvent(this, new StyleChangeEventArgs(_cls, eStyleProperty.IndexedColor, value, _positionID, _address));
		}
	}

	internal override string Id => Theme + Tint + Rgb + Indexed;

	internal ExcelColor(ExcelStyles styles, XmlHelper.ChangedEventHandler ChangedEvent, int worksheetID, string address, eStyleClass cls, StyleBase parent)
		: base(styles, ChangedEvent, worksheetID, address)
	{
		_parent = parent;
		_cls = cls;
	}

	public void SetColor(Color color)
	{
		Rgb = color.ToArgb().ToString("X");
	}

	public void SetColor(int alpha, int red, int green, int blue)
	{
		if (alpha < 0 || red < 0 || green < 0 || blue < 0 || alpha > 255 || red > 255 || green > 255 || blue > 255)
		{
			throw new ArgumentException("Argument range must be from 0 to 255");
		}
		Rgb = alpha.ToString("X2") + red.ToString("X2") + green.ToString("X2") + blue.ToString("X2");
	}

	private ExcelColorXml GetSource()
	{
		base.Index = ((_parent.Index >= 0) ? _parent.Index : 0);
		return _cls switch
		{
			eStyleClass.FillBackgroundColor => _styles.Fills[base.Index].BackgroundColor, 
			eStyleClass.FillPatternColor => _styles.Fills[base.Index].PatternColor, 
			eStyleClass.Font => _styles.Fonts[base.Index].Color, 
			eStyleClass.BorderLeft => _styles.Borders[base.Index].Left.Color, 
			eStyleClass.BorderTop => _styles.Borders[base.Index].Top.Color, 
			eStyleClass.BorderRight => _styles.Borders[base.Index].Right.Color, 
			eStyleClass.BorderBottom => _styles.Borders[base.Index].Bottom.Color, 
			eStyleClass.BorderDiagonal => _styles.Borders[base.Index].Diagonal.Color, 
			_ => throw new Exception("Invalid style-class for Color"), 
		};
	}

	internal override void SetIndex(int index)
	{
		_parent.Index = index;
	}

	public string LookupColor()
	{
		return LookupColor(this);
	}

	public string LookupColor(ExcelColor theColor)
	{
		string text = "";
		string[] array = new string[64]
		{
			"#FF000000", "#FFFFFFFF", "#FFFF0000", "#FF00FF00", "#FF0000FF", "#FFFFFF00", "#FFFF00FF", "#FF00FFFF", "#FF000000", "#FFFFFFFF",
			"#FFFF0000", "#FF00FF00", "#FF0000FF", "#FFFFFF00", "#FFFF00FF", "#FF00FFFF", "#FF800000", "#FF008000", "#FF000080", "#FF808000",
			"#FF800080", "#FF008080", "#FFC0C0C0", "#FF808080", "#FF9999FF", "#FF993366", "#FFFFFFCC", "#FFCCFFFF", "#FF660066", "#FFFF8080",
			"#FF0066CC", "#FFCCCCFF", "#FF000080", "#FFFF00FF", "#FFFFFF00", "#FF00FFFF", "#FF800080", "#FF800000", "#FF008080", "#FF0000FF",
			"#FF00CCFF", "#FFCCFFFF", "#FFCCFFCC", "#FFFFFF99", "#FF99CCFF", "#FFFF99CC", "#FFCC99FF", "#FFFFCC99", "#FF3366FF", "#FF33CCCC",
			"#FF99CC00", "#FFFFCC00", "#FFFF9900", "#FFFF6600", "#FF666699", "#FF969696", "#FF003366", "#FF339966", "#FF003300", "#FF333300",
			"#FF993300", "#FF993366", "#FF333399", "#FF333333"
		};
		if (0 <= theColor.Indexed && array.Length > theColor.Indexed)
		{
			return array[theColor.Indexed];
		}
		if (theColor.Rgb != null && 0 < theColor.Rgb.Length)
		{
			return "#" + theColor.Rgb;
		}
		_ = (int)(theColor.Tint * 160m);
		text = ((int)Math.Round(theColor.Tint * -512m)).ToString("X");
		return "#FF" + text + text + text;
	}
}
