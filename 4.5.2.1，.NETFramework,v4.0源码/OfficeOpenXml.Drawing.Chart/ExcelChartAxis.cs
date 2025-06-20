using System;
using System.Globalization;
using System.Xml;
using OfficeOpenXml.Style;

namespace OfficeOpenXml.Drawing.Chart;

public sealed class ExcelChartAxis : XmlHelper
{
	internal enum eAxisType
	{
		Val,
		Cat,
		Date,
		Serie
	}

	private const string _majorTickMark = "c:majorTickMark/@val";

	private const string _minorTickMark = "c:minorTickMark/@val";

	private string AXIS_POSITION_PATH = "c:axPos/@val";

	private const string _crossesPath = "c:crosses/@val";

	private const string _crossBetweenPath = "c:crossBetween/@val";

	private const string _crossesAtPath = "c:crossesAt/@val";

	private const string _formatPath = "c:numFmt/@formatCode";

	private const string _sourceLinkedPath = "c:numFmt/@sourceLinked";

	private const string _lblPos = "c:tickLblPos/@val";

	private ExcelDrawingFill _fill;

	private ExcelDrawingBorder _border;

	private ExcelTextFont _font;

	private const string _ticLblPos_Path = "c:tickLblPos/@val";

	private const string _displayUnitPath = "c:dispUnits/c:builtInUnit/@val";

	private const string _custUnitPath = "c:dispUnits/c:custUnit/@val";

	private ExcelChartTitle _title;

	private const string _minValuePath = "c:scaling/c:min/@val";

	private const string _maxValuePath = "c:scaling/c:max/@val";

	private const string _majorUnitPath = "c:majorUnit/@val";

	private const string _majorUnitCatPath = "c:tickLblSkip/@val";

	private const string _majorTimeUnitPath = "c:majorTimeUnit/@val";

	private const string _minorUnitPath = "c:minorUnit/@val";

	private const string _minorUnitCatPath = "c:tickMarkSkip/@val";

	private const string _minorTimeUnitPath = "c:minorTimeUnit/@val";

	private const string _logbasePath = "c:scaling/c:logBase/@val";

	private const string _orientationPath = "c:scaling/c:orientation/@val";

	private const string _majorGridlinesPath = "c:majorGridlines";

	private ExcelDrawingBorder _majorGridlines;

	private const string _minorGridlinesPath = "c:minorGridlines";

	private ExcelDrawingBorder _minorGridlines;

	internal string Id => GetXmlNodeString("c:axId/@val");

	public eAxisTickMark MajorTickMark
	{
		get
		{
			string xmlNodeString = GetXmlNodeString("c:majorTickMark/@val");
			if (string.IsNullOrEmpty(xmlNodeString))
			{
				return eAxisTickMark.Cross;
			}
			try
			{
				return (eAxisTickMark)Enum.Parse(typeof(eAxisTickMark), xmlNodeString);
			}
			catch
			{
				return eAxisTickMark.Cross;
			}
		}
		set
		{
			SetXmlNodeString("c:majorTickMark/@val", value.ToString().ToLower(CultureInfo.InvariantCulture));
		}
	}

	public eAxisTickMark MinorTickMark
	{
		get
		{
			string xmlNodeString = GetXmlNodeString("c:minorTickMark/@val");
			if (string.IsNullOrEmpty(xmlNodeString))
			{
				return eAxisTickMark.Cross;
			}
			try
			{
				return (eAxisTickMark)Enum.Parse(typeof(eAxisTickMark), xmlNodeString);
			}
			catch
			{
				return eAxisTickMark.Cross;
			}
		}
		set
		{
			SetXmlNodeString("c:minorTickMark/@val", value.ToString().ToLower(CultureInfo.InvariantCulture));
		}
	}

	internal eAxisType AxisType
	{
		get
		{
			try
			{
				return (eAxisType)Enum.Parse(typeof(eAxisType), base.TopNode.LocalName.Substring(0, 3), ignoreCase: true);
			}
			catch
			{
				return eAxisType.Val;
			}
		}
	}

	public eAxisPosition AxisPosition
	{
		get
		{
			return GetXmlNodeString(AXIS_POSITION_PATH) switch
			{
				"b" => eAxisPosition.Bottom, 
				"r" => eAxisPosition.Right, 
				"t" => eAxisPosition.Top, 
				_ => eAxisPosition.Left, 
			};
		}
		internal set
		{
			SetXmlNodeString(AXIS_POSITION_PATH, value.ToString().ToLower(CultureInfo.InvariantCulture).Substring(0, 1));
		}
	}

	public eCrosses Crosses
	{
		get
		{
			string xmlNodeString = GetXmlNodeString("c:crosses/@val");
			if (string.IsNullOrEmpty(xmlNodeString))
			{
				return eCrosses.AutoZero;
			}
			try
			{
				return (eCrosses)Enum.Parse(typeof(eCrosses), xmlNodeString, ignoreCase: true);
			}
			catch
			{
				return eCrosses.AutoZero;
			}
		}
		set
		{
			string text = value.ToString();
			text = text.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + text.Substring(1, text.Length - 1);
			SetXmlNodeString("c:crosses/@val", text);
		}
	}

	public eCrossBetween CrossBetween
	{
		get
		{
			string xmlNodeString = GetXmlNodeString("c:crossBetween/@val");
			if (string.IsNullOrEmpty(xmlNodeString))
			{
				return eCrossBetween.Between;
			}
			try
			{
				return (eCrossBetween)Enum.Parse(typeof(eCrossBetween), xmlNodeString, ignoreCase: true);
			}
			catch
			{
				return eCrossBetween.Between;
			}
		}
		set
		{
			string text = value.ToString();
			text = text.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + text.Substring(1);
			SetXmlNodeString("c:crossBetween/@val", text);
		}
	}

	public double? CrossesAt
	{
		get
		{
			return GetXmlNodeDoubleNull("c:crossesAt/@val");
		}
		set
		{
			if (!value.HasValue)
			{
				DeleteNode("c:crossesAt/@val");
			}
			else
			{
				SetXmlNodeString("c:crossesAt/@val", value.Value.ToString(CultureInfo.InvariantCulture));
			}
		}
	}

	public string Format
	{
		get
		{
			return GetXmlNodeString("c:numFmt/@formatCode");
		}
		set
		{
			SetXmlNodeString("c:numFmt/@formatCode", value);
			if (string.IsNullOrEmpty(value))
			{
				SourceLinked = true;
			}
			else
			{
				SourceLinked = false;
			}
		}
	}

	public bool SourceLinked
	{
		get
		{
			return GetXmlNodeBool("c:numFmt/@sourceLinked");
		}
		set
		{
			SetXmlNodeBool("c:numFmt/@sourceLinked", value);
		}
	}

	public eTickLabelPosition LabelPosition
	{
		get
		{
			string xmlNodeString = GetXmlNodeString("c:tickLblPos/@val");
			if (string.IsNullOrEmpty(xmlNodeString))
			{
				return eTickLabelPosition.NextTo;
			}
			try
			{
				return (eTickLabelPosition)Enum.Parse(typeof(eTickLabelPosition), xmlNodeString, ignoreCase: true);
			}
			catch
			{
				return eTickLabelPosition.NextTo;
			}
		}
		set
		{
			string text = value.ToString();
			SetXmlNodeString("c:tickLblPos/@val", text.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + text.Substring(1, text.Length - 1));
		}
	}

	public ExcelDrawingFill Fill
	{
		get
		{
			if (_fill == null)
			{
				_fill = new ExcelDrawingFill(base.NameSpaceManager, base.TopNode, "c:spPr");
			}
			return _fill;
		}
	}

	public ExcelDrawingBorder Border
	{
		get
		{
			if (_border == null)
			{
				_border = new ExcelDrawingBorder(base.NameSpaceManager, base.TopNode, "c:spPr/a:ln");
			}
			return _border;
		}
	}

	public ExcelTextFont Font
	{
		get
		{
			if (_font == null)
			{
				if (base.TopNode.SelectSingleNode("c:txPr", base.NameSpaceManager) == null)
				{
					CreateNode("c:txPr/a:bodyPr");
					CreateNode("c:txPr/a:lstStyle");
				}
				_font = new ExcelTextFont(base.NameSpaceManager, base.TopNode, "c:txPr/a:p/a:pPr/a:defRPr", new string[9] { "pPr", "defRPr", "solidFill", "uFill", "latin", "cs", "r", "rPr", "t" });
			}
			return _font;
		}
	}

	public bool Deleted
	{
		get
		{
			return GetXmlNodeBool("c:delete/@val");
		}
		set
		{
			SetXmlNodeBool("c:delete/@val", value);
		}
	}

	public eTickLabelPosition TickLabelPosition
	{
		get
		{
			string xmlNodeString = GetXmlNodeString("c:tickLblPos/@val");
			if (xmlNodeString == "")
			{
				return eTickLabelPosition.None;
			}
			return (eTickLabelPosition)Enum.Parse(typeof(eTickLabelPosition), xmlNodeString, ignoreCase: true);
		}
		set
		{
			string text = value.ToString();
			text = text.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + text.Substring(1, text.Length - 1);
			SetXmlNodeString("c:tickLblPos/@val", text);
		}
	}

	public double DisplayUnit
	{
		get
		{
			string xmlNodeString = GetXmlNodeString("c:dispUnits/c:builtInUnit/@val");
			if (string.IsNullOrEmpty(xmlNodeString))
			{
				double? xmlNodeDoubleNull = GetXmlNodeDoubleNull("c:dispUnits/c:custUnit/@val");
				if (!xmlNodeDoubleNull.HasValue)
				{
					return 0.0;
				}
				return xmlNodeDoubleNull.Value;
			}
			try
			{
				return (long)Enum.Parse(typeof(eBuildInUnits), xmlNodeString, ignoreCase: true);
			}
			catch
			{
				return 0.0;
			}
		}
		set
		{
			if (AxisType != 0 || !(value >= 0.0))
			{
				return;
			}
			foreach (long value2 in Enum.GetValues(typeof(eBuildInUnits)))
			{
				if ((double)value2 == value)
				{
					DeleteNode("c:dispUnits/c:custUnit/@val");
					SetXmlNodeString("c:dispUnits/c:builtInUnit/@val", ((eBuildInUnits)value).ToString());
					return;
				}
			}
			DeleteNode("c:dispUnits/c:builtInUnit/@val");
			if (value != 0.0)
			{
				SetXmlNodeString("c:dispUnits/c:custUnit/@val", value.ToString(CultureInfo.InvariantCulture));
			}
		}
	}

	public ExcelChartTitle Title
	{
		get
		{
			if (_title == null)
			{
				if (base.TopNode.SelectSingleNode("c:title", base.NameSpaceManager) == null)
				{
					CreateNode("c:title");
					base.TopNode.SelectSingleNode("c:title", base.NameSpaceManager).InnerXml = "<c:tx><c:rich><a:bodyPr /><a:lstStyle /><a:p><a:r><a:t /></a:r></a:p></c:rich></c:tx><c:layout /><c:overlay val=\"0\" />";
				}
				_title = new ExcelChartTitle(base.NameSpaceManager, base.TopNode);
			}
			return _title;
		}
	}

	public double? MinValue
	{
		get
		{
			return GetXmlNodeDoubleNull("c:scaling/c:min/@val");
		}
		set
		{
			if (!value.HasValue)
			{
				DeleteNode("c:scaling/c:min/@val");
			}
			else
			{
				SetXmlNodeString("c:scaling/c:min/@val", value.Value.ToString(CultureInfo.InvariantCulture));
			}
		}
	}

	public double? MaxValue
	{
		get
		{
			return GetXmlNodeDoubleNull("c:scaling/c:max/@val");
		}
		set
		{
			if (!value.HasValue)
			{
				DeleteNode("c:scaling/c:max/@val");
			}
			else
			{
				SetXmlNodeString("c:scaling/c:max/@val", value.Value.ToString(CultureInfo.InvariantCulture));
			}
		}
	}

	public double? MajorUnit
	{
		get
		{
			if (AxisType == eAxisType.Cat)
			{
				return GetXmlNodeDoubleNull("c:tickLblSkip/@val");
			}
			return GetXmlNodeDoubleNull("c:majorUnit/@val");
		}
		set
		{
			if (!value.HasValue)
			{
				DeleteNode("c:majorUnit/@val");
				DeleteNode("c:tickLblSkip/@val");
			}
			else if (AxisType == eAxisType.Cat)
			{
				SetXmlNodeString("c:tickLblSkip/@val", value.Value.ToString(CultureInfo.InvariantCulture));
			}
			else
			{
				SetXmlNodeString("c:majorUnit/@val", value.Value.ToString(CultureInfo.InvariantCulture));
			}
		}
	}

	public eTimeUnit? MajorTimeUnit
	{
		get
		{
			return GetXmlNodeString("c:majorTimeUnit/@val") switch
			{
				"years" => eTimeUnit.Years, 
				"months" => eTimeUnit.Months, 
				"days" => eTimeUnit.Days, 
				_ => null, 
			};
		}
		set
		{
			if (value.HasValue)
			{
				SetXmlNodeString("c:majorTimeUnit/@val", value.ToString().ToLower());
			}
			else
			{
				DeleteNode("c:majorTimeUnit/@val");
			}
		}
	}

	public double? MinorUnit
	{
		get
		{
			if (AxisType == eAxisType.Cat)
			{
				return GetXmlNodeDoubleNull("c:tickMarkSkip/@val");
			}
			return GetXmlNodeDoubleNull("c:minorUnit/@val");
		}
		set
		{
			if (!value.HasValue)
			{
				DeleteNode("c:minorUnit/@val");
				DeleteNode("c:tickMarkSkip/@val");
			}
			else if (AxisType == eAxisType.Cat)
			{
				SetXmlNodeString("c:tickMarkSkip/@val", value.Value.ToString(CultureInfo.InvariantCulture));
			}
			else
			{
				SetXmlNodeString("c:minorUnit/@val", value.Value.ToString(CultureInfo.InvariantCulture));
			}
		}
	}

	public eTimeUnit? MinorTimeUnit
	{
		get
		{
			return GetXmlNodeString("c:minorTimeUnit/@val") switch
			{
				"years" => eTimeUnit.Years, 
				"months" => eTimeUnit.Months, 
				"days" => eTimeUnit.Days, 
				_ => null, 
			};
		}
		set
		{
			if (value.HasValue)
			{
				SetXmlNodeString("c:minorTimeUnit/@val", value.ToString().ToLower());
			}
			else
			{
				DeleteNode("c:minorTimeUnit/@val");
			}
		}
	}

	public double? LogBase
	{
		get
		{
			return GetXmlNodeDoubleNull("c:scaling/c:logBase/@val");
		}
		set
		{
			if (!value.HasValue)
			{
				DeleteNode("c:scaling/c:logBase/@val");
				return;
			}
			double num = value.Value;
			if (num < 2.0 || num > 1000.0)
			{
				throw new ArgumentOutOfRangeException("Value must be between 2 and 1000");
			}
			SetXmlNodeString("c:scaling/c:logBase/@val", num.ToString("0.0", CultureInfo.InvariantCulture));
		}
	}

	public eAxisOrientation Orientation
	{
		get
		{
			string xmlNodeString = GetXmlNodeString("c:scaling/c:orientation/@val");
			if (xmlNodeString == "")
			{
				return eAxisOrientation.MinMax;
			}
			return (eAxisOrientation)Enum.Parse(typeof(eAxisOrientation), xmlNodeString, ignoreCase: true);
		}
		set
		{
			string text = value.ToString();
			text = text.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + text.Substring(1, text.Length - 1);
			SetXmlNodeString("c:scaling/c:orientation/@val", text);
		}
	}

	public ExcelDrawingBorder MajorGridlines
	{
		get
		{
			if (_majorGridlines == null)
			{
				if (base.TopNode.SelectSingleNode("c:majorGridlines", base.NameSpaceManager) == null)
				{
					CreateNode("c:majorGridlines");
				}
				_majorGridlines = new ExcelDrawingBorder(base.NameSpaceManager, base.TopNode, string.Format("{0}/c:spPr/a:ln", "c:majorGridlines"));
			}
			return _majorGridlines;
		}
	}

	public ExcelDrawingBorder MinorGridlines
	{
		get
		{
			if (_minorGridlines == null)
			{
				if (base.TopNode.SelectSingleNode("c:minorGridlines", base.NameSpaceManager) == null)
				{
					CreateNode("c:minorGridlines");
				}
				_minorGridlines = new ExcelDrawingBorder(base.NameSpaceManager, base.TopNode, string.Format("{0}/c:spPr/a:ln", "c:minorGridlines"));
			}
			return _minorGridlines;
		}
	}

	internal ExcelChartAxis(XmlNamespaceManager nameSpaceManager, XmlNode topNode)
		: base(nameSpaceManager, topNode)
	{
		base.SchemaNodeOrder = new string[30]
		{
			"axId", "scaling", "logBase", "orientation", "max", "min", "delete", "axPos", "majorGridlines", "minorGridlines",
			"title", "numFmt", "majorTickMark", "minorTickMark", "tickLblPos", "spPr", "txPr", "crossAx", "crossesAt", "crosses",
			"crossBetween", "auto", "lblOffset", "majorUnit", "majorTimeUnit", "minorUnit", "minorTimeUnit", "dispUnits", "spPr", "txPr"
		};
	}

	public void RemoveGridlines()
	{
		RemoveGridlines(removeMajor: true, removeMinor: true);
	}

	public void RemoveGridlines(bool removeMajor, bool removeMinor)
	{
		if (removeMajor)
		{
			DeleteNode("c:majorGridlines");
			_majorGridlines = null;
		}
		if (removeMinor)
		{
			DeleteNode("c:minorGridlines");
			_minorGridlines = null;
		}
	}
}
