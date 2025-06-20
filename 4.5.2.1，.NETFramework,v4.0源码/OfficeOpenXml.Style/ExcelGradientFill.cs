using System;
using OfficeOpenXml.Style.XmlAccess;

namespace OfficeOpenXml.Style;

public class ExcelGradientFill : StyleBase
{
	private ExcelColor _gradientColor1;

	private ExcelColor _gradientColor2;

	public double Degree
	{
		get
		{
			return ((ExcelGradientFillXml)_styles.Fills[base.Index]).Degree;
		}
		set
		{
			_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.GradientFill, eStyleProperty.GradientDegree, value, _positionID, _address));
		}
	}

	public ExcelFillGradientType Type
	{
		get
		{
			return ((ExcelGradientFillXml)_styles.Fills[base.Index]).Type;
		}
		set
		{
			_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.GradientFill, eStyleProperty.GradientType, value, _positionID, _address));
		}
	}

	public double Top
	{
		get
		{
			return ((ExcelGradientFillXml)_styles.Fills[base.Index]).Top;
		}
		set
		{
			if (value < 0.0 || value > 1.0)
			{
				throw new ArgumentOutOfRangeException("Value must be between 0 and 1");
			}
			_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.GradientFill, eStyleProperty.GradientTop, value, _positionID, _address));
		}
	}

	public double Bottom
	{
		get
		{
			return ((ExcelGradientFillXml)_styles.Fills[base.Index]).Bottom;
		}
		set
		{
			if (value < 0.0 || value > 1.0)
			{
				throw new ArgumentOutOfRangeException("Value must be between 0 and 1");
			}
			_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.GradientFill, eStyleProperty.GradientBottom, value, _positionID, _address));
		}
	}

	public double Left
	{
		get
		{
			return ((ExcelGradientFillXml)_styles.Fills[base.Index]).Left;
		}
		set
		{
			if (value < 0.0 || value > 1.0)
			{
				throw new ArgumentOutOfRangeException("Value must be between 0 and 1");
			}
			_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.GradientFill, eStyleProperty.GradientLeft, value, _positionID, _address));
		}
	}

	public double Right
	{
		get
		{
			return ((ExcelGradientFillXml)_styles.Fills[base.Index]).Right;
		}
		set
		{
			if (value < 0.0 || value > 1.0)
			{
				throw new ArgumentOutOfRangeException("Value must be between 0 and 1");
			}
			_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.GradientFill, eStyleProperty.GradientRight, value, _positionID, _address));
		}
	}

	public ExcelColor Color1
	{
		get
		{
			if (_gradientColor1 == null)
			{
				_gradientColor1 = new ExcelColor(_styles, _ChangedEvent, _positionID, _address, eStyleClass.FillGradientColor1, this);
			}
			return _gradientColor1;
		}
	}

	public ExcelColor Color2
	{
		get
		{
			if (_gradientColor2 == null)
			{
				_gradientColor2 = new ExcelColor(_styles, _ChangedEvent, _positionID, _address, eStyleClass.FillGradientColor2, this);
			}
			return _gradientColor2;
		}
	}

	internal override string Id => string.Concat(Degree.ToString(), Type, Color1.Id, Color2.Id, Top.ToString(), Bottom.ToString(), Left.ToString(), Right.ToString());

	internal ExcelGradientFill(ExcelStyles styles, XmlHelper.ChangedEventHandler ChangedEvent, int PositionID, string address, int index)
		: base(styles, ChangedEvent, PositionID, address)
	{
		base.Index = index;
	}
}
