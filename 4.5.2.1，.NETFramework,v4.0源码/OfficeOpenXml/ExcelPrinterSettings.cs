using System;
using System.Globalization;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml;

public sealed class ExcelPrinterSettings : XmlHelper
{
	private ExcelWorksheet _ws;

	private bool _marginsCreated;

	private const string _leftMarginPath = "d:pageMargins/@left";

	private const string _rightMarginPath = "d:pageMargins/@right";

	private const string _topMarginPath = "d:pageMargins/@top";

	private const string _bottomMarginPath = "d:pageMargins/@bottom";

	private const string _headerMarginPath = "d:pageMargins/@header";

	private const string _footerMarginPath = "d:pageMargins/@footer";

	private const string _orientationPath = "d:pageSetup/@orientation";

	private const string _fitToWidthPath = "d:pageSetup/@fitToWidth";

	private const string _fitToHeightPath = "d:pageSetup/@fitToHeight";

	private const string _scalePath = "d:pageSetup/@scale";

	private const string _fitToPagePath = "d:sheetPr/d:pageSetUpPr/@fitToPage";

	private const string _headersPath = "d:printOptions/@headings";

	private const string _gridLinesPath = "d:printOptions/@gridLines";

	private const string _horizontalCenteredPath = "d:printOptions/@horizontalCentered";

	private const string _verticalCenteredPath = "d:printOptions/@verticalCentered";

	private const string _pageOrderPath = "d:pageSetup/@pageOrder";

	private const string _blackAndWhitePath = "d:pageSetup/@blackAndWhite";

	private const string _draftPath = "d:pageSetup/@draft";

	private const string _paperSizePath = "d:pageSetup/@paperSize";

	public decimal LeftMargin
	{
		get
		{
			return GetXmlNodeDecimal("d:pageMargins/@left");
		}
		set
		{
			CreateMargins();
			SetXmlNodeString("d:pageMargins/@left", value.ToString(CultureInfo.InvariantCulture));
		}
	}

	public decimal RightMargin
	{
		get
		{
			return GetXmlNodeDecimal("d:pageMargins/@right");
		}
		set
		{
			CreateMargins();
			SetXmlNodeString("d:pageMargins/@right", value.ToString(CultureInfo.InvariantCulture));
		}
	}

	public decimal TopMargin
	{
		get
		{
			return GetXmlNodeDecimal("d:pageMargins/@top");
		}
		set
		{
			CreateMargins();
			SetXmlNodeString("d:pageMargins/@top", value.ToString(CultureInfo.InvariantCulture));
		}
	}

	public decimal BottomMargin
	{
		get
		{
			return GetXmlNodeDecimal("d:pageMargins/@bottom");
		}
		set
		{
			CreateMargins();
			SetXmlNodeString("d:pageMargins/@bottom", value.ToString(CultureInfo.InvariantCulture));
		}
	}

	public decimal HeaderMargin
	{
		get
		{
			return GetXmlNodeDecimal("d:pageMargins/@header");
		}
		set
		{
			CreateMargins();
			SetXmlNodeString("d:pageMargins/@header", value.ToString(CultureInfo.InvariantCulture));
		}
	}

	public decimal FooterMargin
	{
		get
		{
			return GetXmlNodeDecimal("d:pageMargins/@footer");
		}
		set
		{
			CreateMargins();
			SetXmlNodeString("d:pageMargins/@footer", value.ToString(CultureInfo.InvariantCulture));
		}
	}

	public eOrientation Orientation
	{
		get
		{
			return (eOrientation)Enum.Parse(typeof(eOrientation), GetXmlNodeString("d:pageSetup/@orientation"), ignoreCase: true);
		}
		set
		{
			SetXmlNodeString("d:pageSetup/@orientation", value.ToString().ToLower(CultureInfo.InvariantCulture));
		}
	}

	public int FitToWidth
	{
		get
		{
			return GetXmlNodeInt("d:pageSetup/@fitToWidth");
		}
		set
		{
			SetXmlNodeString("d:pageSetup/@fitToWidth", value.ToString());
		}
	}

	public int FitToHeight
	{
		get
		{
			return GetXmlNodeInt("d:pageSetup/@fitToHeight");
		}
		set
		{
			SetXmlNodeString("d:pageSetup/@fitToHeight", value.ToString());
		}
	}

	public int Scale
	{
		get
		{
			return GetXmlNodeInt("d:pageSetup/@scale");
		}
		set
		{
			SetXmlNodeString("d:pageSetup/@scale", value.ToString());
		}
	}

	public bool FitToPage
	{
		get
		{
			return GetXmlNodeBool("d:sheetPr/d:pageSetUpPr/@fitToPage");
		}
		set
		{
			SetXmlNodeString("d:sheetPr/d:pageSetUpPr/@fitToPage", value ? "1" : "0");
		}
	}

	public bool ShowHeaders
	{
		get
		{
			return GetXmlNodeBool("d:printOptions/@headings", blankValue: false);
		}
		set
		{
			SetXmlNodeBool("d:printOptions/@headings", value, removeIf: false);
		}
	}

	public ExcelAddress RepeatRows
	{
		get
		{
			if (_ws.Names.ContainsKey("_xlnm.Print_Titles"))
			{
				ExcelRangeBase excelRangeBase = _ws.Names["_xlnm.Print_Titles"];
				if (excelRangeBase.Start.Column == 1 && excelRangeBase.End.Column == 16384)
				{
					return new ExcelAddress(excelRangeBase.FirstAddress);
				}
				if (excelRangeBase._addresses != null)
				{
					return excelRangeBase._addresses.FirstOrDefault((ExcelAddress a) => a.Start.Column == 1 && a.End.Column == 16384);
				}
				return null;
			}
			return null;
		}
		set
		{
			if (value.Start.Column != 1 || value.End.Column != 16384)
			{
				throw new InvalidOperationException("Address must span full columns only (for ex. Address=\"A:A\" for the first column).");
			}
			ExcelAddress repeatColumns = RepeatColumns;
			string address = ((repeatColumns != null) ? (repeatColumns.Address + "," + value.Address) : value.Address);
			if (_ws.Names.ContainsKey("_xlnm.Print_Titles"))
			{
				_ws.Names["_xlnm.Print_Titles"].Address = address;
			}
			else
			{
				_ws.Names.Add("_xlnm.Print_Titles", new ExcelRangeBase(_ws, address));
			}
		}
	}

	public ExcelAddress RepeatColumns
	{
		get
		{
			if (_ws.Names.ContainsKey("_xlnm.Print_Titles"))
			{
				ExcelRangeBase excelRangeBase = _ws.Names["_xlnm.Print_Titles"];
				if (excelRangeBase.Start.Row == 1 && excelRangeBase.End.Row == 1048576)
				{
					return new ExcelAddress(excelRangeBase.FirstAddress);
				}
				if (excelRangeBase._addresses != null)
				{
					return excelRangeBase._addresses.FirstOrDefault((ExcelAddress a) => a.Start.Row == 1 && a.End.Row == 1048576);
				}
				return null;
			}
			return null;
		}
		set
		{
			if (value.Start.Row != 1 || value.End.Row != 1048576)
			{
				throw new InvalidOperationException("Address must span rows only (for ex. Address=\"1:1\" for the first row).");
			}
			ExcelAddress repeatRows = RepeatRows;
			string address = ((repeatRows != null) ? (value.Address + "," + repeatRows.Address) : value.Address);
			if (_ws.Names.ContainsKey("_xlnm.Print_Titles"))
			{
				_ws.Names["_xlnm.Print_Titles"].Address = address;
			}
			else
			{
				_ws.Names.Add("_xlnm.Print_Titles", new ExcelRangeBase(_ws, address));
			}
		}
	}

	public ExcelRangeBase PrintArea
	{
		get
		{
			if (_ws.Names.ContainsKey("_xlnm.Print_Area"))
			{
				return _ws.Names["_xlnm.Print_Area"];
			}
			return null;
		}
		set
		{
			if (value == null)
			{
				_ws.Names.Remove("_xlnm.Print_Area");
			}
			else if (_ws.Names.ContainsKey("_xlnm.Print_Area"))
			{
				_ws.Names["_xlnm.Print_Area"].Address = value.Address;
			}
			else
			{
				_ws.Names.Add("_xlnm.Print_Area", value);
			}
		}
	}

	public bool ShowGridLines
	{
		get
		{
			return GetXmlNodeBool("d:printOptions/@gridLines", blankValue: false);
		}
		set
		{
			SetXmlNodeBool("d:printOptions/@gridLines", value, removeIf: false);
		}
	}

	public bool HorizontalCentered
	{
		get
		{
			return GetXmlNodeBool("d:printOptions/@horizontalCentered", blankValue: false);
		}
		set
		{
			SetXmlNodeBool("d:printOptions/@horizontalCentered", value, removeIf: false);
		}
	}

	public bool VerticalCentered
	{
		get
		{
			return GetXmlNodeBool("d:printOptions/@verticalCentered", blankValue: false);
		}
		set
		{
			SetXmlNodeBool("d:printOptions/@verticalCentered", value, removeIf: false);
		}
	}

	public ePageOrder PageOrder
	{
		get
		{
			if (GetXmlNodeString("d:pageSetup/@pageOrder") == "overThenDown")
			{
				return ePageOrder.OverThenDown;
			}
			return ePageOrder.DownThenOver;
		}
		set
		{
			if (value == ePageOrder.OverThenDown)
			{
				SetXmlNodeString("d:pageSetup/@pageOrder", "overThenDown");
			}
			else
			{
				DeleteNode("d:pageSetup/@pageOrder");
			}
		}
	}

	public bool BlackAndWhite
	{
		get
		{
			return GetXmlNodeBool("d:pageSetup/@blackAndWhite", blankValue: false);
		}
		set
		{
			SetXmlNodeBool("d:pageSetup/@blackAndWhite", value, removeIf: false);
		}
	}

	public bool Draft
	{
		get
		{
			return GetXmlNodeBool("d:pageSetup/@draft", blankValue: false);
		}
		set
		{
			SetXmlNodeBool("d:pageSetup/@draft", value, removeIf: false);
		}
	}

	public ePaperSize PaperSize
	{
		get
		{
			string xmlNodeString = GetXmlNodeString("d:pageSetup/@paperSize");
			if (xmlNodeString != "")
			{
				return (ePaperSize)int.Parse(xmlNodeString);
			}
			return ePaperSize.Letter;
		}
		set
		{
			int num = (int)value;
			SetXmlNodeString("d:pageSetup/@paperSize", num.ToString());
		}
	}

	internal ExcelPrinterSettings(XmlNamespaceManager ns, XmlNode topNode, ExcelWorksheet ws)
		: base(ns, topNode)
	{
		_ws = ws;
		base.SchemaNodeOrder = ws.SchemaNodeOrder;
	}

	private void CreateMargins()
	{
		if (!_marginsCreated && base.TopNode.SelectSingleNode("d:pageMargins/@left", base.NameSpaceManager) == null)
		{
			_marginsCreated = true;
			LeftMargin = 0.7087m;
			RightMargin = 0.7087m;
			TopMargin = 0.7480m;
			BottomMargin = 0.7480m;
			HeaderMargin = 0.315m;
			FooterMargin = 0.315m;
		}
	}
}
