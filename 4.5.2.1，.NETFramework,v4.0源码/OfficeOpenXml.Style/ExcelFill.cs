namespace OfficeOpenXml.Style;

public class ExcelFill : StyleBase
{
	private ExcelColor _patternColor;

	private ExcelColor _backgroundColor;

	private ExcelGradientFill _gradient;

	public ExcelFillStyle PatternType
	{
		get
		{
			if (base.Index == int.MinValue)
			{
				return ExcelFillStyle.None;
			}
			return _styles.Fills[base.Index].PatternType;
		}
		set
		{
			if (_gradient != null)
			{
				_gradient = null;
			}
			_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Fill, eStyleProperty.PatternType, value, _positionID, _address));
		}
	}

	public ExcelColor PatternColor
	{
		get
		{
			if (_patternColor == null)
			{
				_patternColor = new ExcelColor(_styles, _ChangedEvent, _positionID, _address, eStyleClass.FillPatternColor, this);
				if (_gradient != null)
				{
					_gradient = null;
				}
			}
			return _patternColor;
		}
	}

	public ExcelColor BackgroundColor
	{
		get
		{
			if (_backgroundColor == null)
			{
				_backgroundColor = new ExcelColor(_styles, _ChangedEvent, _positionID, _address, eStyleClass.FillBackgroundColor, this);
				if (_gradient != null)
				{
					_gradient = null;
				}
			}
			return _backgroundColor;
		}
	}

	public ExcelGradientFill Gradient
	{
		get
		{
			if (_gradient == null)
			{
				_gradient = new ExcelGradientFill(_styles, _ChangedEvent, _positionID, _address, base.Index);
				_backgroundColor = null;
				_patternColor = null;
			}
			return _gradient;
		}
	}

	internal override string Id
	{
		get
		{
			if (_gradient == null)
			{
				return string.Concat(PatternType, PatternColor.Id, BackgroundColor.Id);
			}
			return _gradient.Id;
		}
	}

	internal ExcelFill(ExcelStyles styles, XmlHelper.ChangedEventHandler ChangedEvent, int PositionID, string address, int index)
		: base(styles, ChangedEvent, PositionID, address)
	{
		base.Index = index;
	}
}
