using System;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Sparkline;

public class ExcelSparklineGroup : XmlHelper
{
	private ExcelWorksheet _ws;

	private const string _dateAxisPath = "@dateAxis";

	private const string _markersPath = "@markers";

	private const string _highPath = "@high";

	private const string _lowPath = "@low";

	private const string _firstPath = "@first";

	private const string _lastPath = "@last";

	private const string _negativePath = "@negative";

	private const string _displayXAxisPath = "@displayXAxis";

	private const string _displayHiddenPath = "@displayHidden";

	private const string lineWidthPath = "x14:sparklineGroup/@lineWidth";

	private const string _dispBlanksAsPath = "@displayEmptyCellsAs";

	private const string _typePath = "@type";

	private const string _colorSeriesPath = "x14:colorSeries";

	private const string _colorNegativePath = "x14:colorNegative";

	private const string _colorAxisPath = "x14:colorAxis";

	private const string _colorMarkersPath = "x14:colorMarkers";

	private const string _colorFirstPath = "x14:colorFirst";

	private const string _colorLastPath = "x14:colorLast";

	private const string _colorHighPath = "x14:colorHigh";

	private const string _colorLowPath = "x14:colorLow";

	private const string _manualMinPath = "@manualMin";

	private const string _manualMaxPath = "@manualMax";

	private const string _minAxisTypePath = "@minAxisType";

	private const string _maxAxisTypePath = "@maxAxisType";

	private const string _rightToLeftPath = "@rightToLeft";

	public ExcelRangeBase DateAxisRange
	{
		get
		{
			string xmlNodeString = GetXmlNodeString("xm:f");
			if (string.IsNullOrEmpty(xmlNodeString))
			{
				return null;
			}
			ExcelAddressBase excelAddressBase = new ExcelAddressBase(xmlNodeString);
			if (excelAddressBase.WorkSheet.Equals(_ws.Name, StringComparison.CurrentCultureIgnoreCase))
			{
				return _ws.Cells[excelAddressBase.Address];
			}
			return _ws.Workbook.Worksheets[excelAddressBase.WorkSheet].Cells[excelAddressBase.Address];
		}
		set
		{
			if (value == null)
			{
				RemoveDateAxis();
				return;
			}
			if (value.Worksheet.Workbook != _ws.Workbook)
			{
				throw new ArgumentException("Range must be in the same package");
			}
			if (value.Rows != 1 && value.Columns != 1)
			{
				throw new ArgumentException("Range must only be 1 row or column");
			}
			DateAxis = true;
			SetXmlNodeString("xm:f", value.FullAddress);
		}
	}

	public ExcelRangeBase DataRange
	{
		get
		{
			if (Sparklines.Count == 0)
			{
				return null;
			}
			_ = _ws.Workbook.Worksheets[Sparklines[0].RangeAddress.WorkSheet];
			return _ws.Cells[Sparklines[0].RangeAddress._fromRow, Sparklines[0].RangeAddress._fromCol, Sparklines[Sparklines.Count - 1].RangeAddress._toRow, Sparklines[Sparklines.Count - 1].RangeAddress._toCol];
		}
	}

	public ExcelRangeBase LocationRange
	{
		get
		{
			if (Sparklines.Count == 0)
			{
				return null;
			}
			return _ws.Cells[Sparklines[0].Cell.Row, Sparklines[0].Cell.Column, Sparklines[Sparklines.Count - 1].Cell.Row, Sparklines[Sparklines.Count - 1].Cell.Column];
		}
	}

	public ExcelSparklineCollection Sparklines { get; internal set; }

	internal bool DateAxis
	{
		get
		{
			return GetXmlNodeBool("@dateAxis", blankValue: false);
		}
		set
		{
			SetXmlNodeBool("@dateAxis", value);
		}
	}

	public bool Markers
	{
		get
		{
			return GetXmlNodeBool("@markers", blankValue: false);
		}
		set
		{
			SetXmlNodeBool("@markers", value);
		}
	}

	public bool High
	{
		get
		{
			return GetXmlNodeBool("@high", blankValue: false);
		}
		set
		{
			SetXmlNodeBool("@high", value);
		}
	}

	public bool Low
	{
		get
		{
			return GetXmlNodeBool("@low", blankValue: false);
		}
		set
		{
			SetXmlNodeBool("@low", value);
		}
	}

	public bool First
	{
		get
		{
			return GetXmlNodeBool("@first", blankValue: false);
		}
		set
		{
			SetXmlNodeBool("@first", value);
		}
	}

	public bool Last
	{
		get
		{
			return GetXmlNodeBool("@last", blankValue: false);
		}
		set
		{
			SetXmlNodeBool("@last", value);
		}
	}

	public bool Negative
	{
		get
		{
			return GetXmlNodeBool("@negative");
		}
		set
		{
			SetXmlNodeBool("@negative", value);
		}
	}

	public bool DisplayXAxis
	{
		get
		{
			return GetXmlNodeBool("@displayXAxis");
		}
		set
		{
			SetXmlNodeBool("@displayXAxis", value);
		}
	}

	public bool DisplayHidden
	{
		get
		{
			return GetXmlNodeBool("@displayHidden");
		}
		set
		{
			SetXmlNodeBool("@displayHidden", value);
		}
	}

	public double LineWidth
	{
		get
		{
			return GetXmlNodeDoubleNull("x14:sparklineGroup/@lineWidth") ?? 0.75;
		}
		set
		{
			SetXmlNodeString("x14:sparklineGroup/@lineWidth", value.ToString(CultureInfo.InvariantCulture));
		}
	}

	public eDispBlanksAs DisplayEmptyCellsAs
	{
		get
		{
			string xmlNodeString = GetXmlNodeString("@displayEmptyCellsAs");
			if (string.IsNullOrEmpty(xmlNodeString))
			{
				return eDispBlanksAs.Zero;
			}
			return (eDispBlanksAs)Enum.Parse(typeof(eDispBlanksAs), xmlNodeString, ignoreCase: true);
		}
		set
		{
			SetXmlNodeString("@displayEmptyCellsAs", value.ToString().ToLower());
		}
	}

	public eSparklineType Type
	{
		get
		{
			string xmlNodeString = GetXmlNodeString("@type");
			if (string.IsNullOrEmpty(xmlNodeString))
			{
				return eSparklineType.Line;
			}
			return (eSparklineType)Enum.Parse(typeof(eSparklineType), xmlNodeString, ignoreCase: true);
		}
		set
		{
			SetXmlNodeString("@type", value.ToString().ToLower());
		}
	}

	public ExcelSparklineColor ColorSeries
	{
		get
		{
			CreateNode("x14:colorSeries");
			return new ExcelSparklineColor(base.NameSpaceManager, base.TopNode.SelectSingleNode("x14:colorSeries", base.NameSpaceManager));
		}
	}

	public ExcelSparklineColor ColorNegative
	{
		get
		{
			CreateNode("x14:colorNegative");
			return new ExcelSparklineColor(base.NameSpaceManager, base.TopNode.SelectSingleNode("x14:colorNegative", base.NameSpaceManager));
		}
	}

	public ExcelSparklineColor ColorAxis
	{
		get
		{
			CreateNode("x14:colorAxis");
			return new ExcelSparklineColor(base.NameSpaceManager, base.TopNode.SelectSingleNode("x14:colorAxis", base.NameSpaceManager));
		}
	}

	public ExcelSparklineColor ColorMarkers
	{
		get
		{
			CreateNode("x14:colorMarkers");
			return new ExcelSparklineColor(base.NameSpaceManager, base.TopNode.SelectSingleNode("x14:colorMarkers", base.NameSpaceManager));
		}
	}

	public ExcelSparklineColor ColorFirst
	{
		get
		{
			CreateNode("x14:colorFirst");
			return new ExcelSparklineColor(base.NameSpaceManager, base.TopNode.SelectSingleNode("x14:colorFirst", base.NameSpaceManager));
		}
	}

	public ExcelSparklineColor ColorLast
	{
		get
		{
			CreateNode("x14:colorLast");
			return new ExcelSparklineColor(base.NameSpaceManager, base.TopNode.SelectSingleNode("x14:colorLast", base.NameSpaceManager));
		}
	}

	public ExcelSparklineColor ColorHigh
	{
		get
		{
			CreateNode("x14:colorHigh");
			return new ExcelSparklineColor(base.NameSpaceManager, base.TopNode.SelectSingleNode("x14:colorHigh", base.NameSpaceManager));
		}
	}

	public ExcelSparklineColor ColorLow
	{
		get
		{
			CreateNode("x14:colorLow");
			return new ExcelSparklineColor(base.NameSpaceManager, base.TopNode.SelectSingleNode("x14:colorLow", base.NameSpaceManager));
		}
	}

	public double ManualMin
	{
		get
		{
			return GetXmlNodeDouble("@manualMin");
		}
		set
		{
			SetXmlNodeString("@minAxisType", "custom");
			SetXmlNodeString("@manualMin", value.ToString("F", CultureInfo.InvariantCulture));
		}
	}

	public double ManualMax
	{
		get
		{
			return GetXmlNodeDouble("@manualMax");
		}
		set
		{
			SetXmlNodeString("@maxAxisType", "custom");
			SetXmlNodeString("@manualMax", value.ToString("F", CultureInfo.InvariantCulture));
		}
	}

	public eSparklineAxisMinMax MinAxisType
	{
		get
		{
			string xmlNodeString = GetXmlNodeString("@minAxisType");
			if (string.IsNullOrEmpty(xmlNodeString))
			{
				return eSparklineAxisMinMax.Individual;
			}
			return (eSparklineAxisMinMax)Enum.Parse(typeof(eSparklineAxisMinMax), xmlNodeString, ignoreCase: true);
		}
		set
		{
			if (value == eSparklineAxisMinMax.Custom)
			{
				ManualMin = 0.0;
				return;
			}
			SetXmlNodeString("@minAxisType", value.ToString());
			DeleteNode("@manualMin");
		}
	}

	public eSparklineAxisMinMax MaxAxisType
	{
		get
		{
			string xmlNodeString = GetXmlNodeString("@maxAxisType");
			if (string.IsNullOrEmpty(xmlNodeString))
			{
				return eSparklineAxisMinMax.Individual;
			}
			return (eSparklineAxisMinMax)Enum.Parse(typeof(eSparklineAxisMinMax), xmlNodeString, ignoreCase: true);
		}
		set
		{
			if (value == eSparklineAxisMinMax.Custom)
			{
				ManualMax = 0.0;
				return;
			}
			SetXmlNodeString("@maxAxisType", value.ToString());
			DeleteNode("@manualMax");
		}
	}

	public bool RightToLeft
	{
		get
		{
			return GetXmlNodeBool("@rightToLeft", blankValue: false);
		}
		set
		{
			SetXmlNodeBool("@rightToLeft", value);
		}
	}

	internal ExcelSparklineGroup(XmlNamespaceManager ns, XmlElement topNode, ExcelWorksheet ws)
		: base(ns, topNode)
	{
		base.SchemaNodeOrder = new string[10] { "colorSeries", "colorNegative", "colorAxis", "colorMarkers", "colorFirst", "colorLast", "colorHigh", "colorLow", "f", "sparklines" };
		Sparklines = new ExcelSparklineCollection(this);
		_ws = ws;
	}

	private void RemoveDateAxis()
	{
		DeleteNode("xm:f");
		DateAxis = false;
	}
}
