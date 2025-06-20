using System;
using System.Globalization;
using System.Xml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style.XmlAccess;

namespace OfficeOpenXml.Drawing;

public class ExcelDrawing : XmlHelper, IDisposable
{
	public class ExcelPosition : XmlHelper
	{
		internal delegate void SetWidthCallback();

		private XmlNode _node;

		private XmlNamespaceManager _ns;

		private SetWidthCallback _setWidthCallback;

		private const string colPath = "xdr:col";

		private const string rowPath = "xdr:row";

		private const string colOffPath = "xdr:colOff";

		private const string rowOffPath = "xdr:rowOff";

		public int Column
		{
			get
			{
				return GetXmlNodeInt("xdr:col");
			}
			set
			{
				SetXmlNodeString("xdr:col", value.ToString());
				_setWidthCallback();
			}
		}

		public int Row
		{
			get
			{
				return GetXmlNodeInt("xdr:row");
			}
			set
			{
				SetXmlNodeString("xdr:row", value.ToString());
				_setWidthCallback();
			}
		}

		public int ColumnOff
		{
			get
			{
				return GetXmlNodeInt("xdr:colOff");
			}
			set
			{
				SetXmlNodeString("xdr:colOff", value.ToString());
				_setWidthCallback();
			}
		}

		public int RowOff
		{
			get
			{
				return GetXmlNodeInt("xdr:rowOff");
			}
			set
			{
				SetXmlNodeString("xdr:rowOff", value.ToString());
				_setWidthCallback();
			}
		}

		internal ExcelPosition(XmlNamespaceManager ns, XmlNode node, SetWidthCallback setWidthCallback)
			: base(ns, node)
		{
			_node = node;
			_ns = ns;
			_setWidthCallback = setWidthCallback;
		}
	}

	protected ExcelDrawings _drawings;

	protected XmlNode _topNode;

	private string _nameXPath;

	protected internal int _id;

	private const float STANDARD_DPI = 96f;

	public const int EMU_PER_PIXEL = 9525;

	protected internal int _width = int.MinValue;

	protected internal int _height = int.MinValue;

	protected internal int _top = int.MinValue;

	protected internal int _left = int.MinValue;

	private bool _doNotAdjust;

	private const string lockedPath = "xdr:clientData/@fLocksWithSheet";

	private const string printPath = "xdr:clientData/@fPrintsWithSheet";

	private ExcelPosition _to;

	public string Name
	{
		get
		{
			try
			{
				if (_nameXPath == "")
				{
					return "";
				}
				return GetXmlNodeString(_nameXPath);
			}
			catch
			{
				return "";
			}
		}
		set
		{
			try
			{
				if (_nameXPath == "")
				{
					throw new NotImplementedException();
				}
				SetXmlNodeString(_nameXPath, value);
			}
			catch
			{
				throw new NotImplementedException();
			}
		}
	}

	public eEditAs EditAs
	{
		get
		{
			try
			{
				string xmlNodeString = GetXmlNodeString("@editAs");
				if (xmlNodeString == "")
				{
					return eEditAs.TwoCell;
				}
				return (eEditAs)Enum.Parse(typeof(eEditAs), xmlNodeString, ignoreCase: true);
			}
			catch
			{
				return eEditAs.TwoCell;
			}
		}
		set
		{
			string text = value.ToString();
			SetXmlNodeString("@editAs", text.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + text.Substring(1, text.Length - 1));
		}
	}

	public bool Locked
	{
		get
		{
			return GetXmlNodeBool("xdr:clientData/@fLocksWithSheet", blankValue: true);
		}
		set
		{
			SetXmlNodeBool("xdr:clientData/@fLocksWithSheet", value);
		}
	}

	public bool Print
	{
		get
		{
			return GetXmlNodeBool("xdr:clientData/@fPrintsWithSheet", blankValue: true);
		}
		set
		{
			SetXmlNodeBool("xdr:clientData/@fPrintsWithSheet", value);
		}
	}

	public ExcelPosition From { get; private set; }

	public ExcelPosition To
	{
		get
		{
			return _to;
		}
		private set
		{
			_to = value;
		}
	}

	internal string Id => _id.ToString();

	internal ExcelDrawing(ExcelDrawings drawings, XmlNode node, string nameXPath)
		: base(drawings.NameSpaceManager, node)
	{
		_drawings = drawings;
		_topNode = node;
		_id = drawings.Worksheet.Workbook._nextDrawingID++;
		XmlNode node2 = node.SelectSingleNode("xdr:from", drawings.NameSpaceManager);
		if (node != null)
		{
			From = new ExcelPosition(drawings.NameSpaceManager, node2, GetPositionSize);
		}
		node2 = node.SelectSingleNode("xdr:to", drawings.NameSpaceManager);
		if (node != null)
		{
			To = new ExcelPosition(drawings.NameSpaceManager, node2, GetPositionSize);
		}
		else
		{
			To = null;
		}
		GetPositionSize();
		_nameXPath = nameXPath;
		base.SchemaNodeOrder = new string[5] { "from", "to", "graphicFrame", "sp", "clientData" };
	}

	internal void ReSetWidthHeight()
	{
		if (!_drawings.Worksheet.Workbook._package.DoAdjustDrawings && EditAs == eEditAs.Absolute)
		{
			SetPixelWidth(_width);
			SetPixelHeight(_height);
		}
	}

	internal static ExcelDrawing GetDrawing(ExcelDrawings drawings, XmlNode node)
	{
		if (node.SelectSingleNode("xdr:sp", drawings.NameSpaceManager) != null)
		{
			return new ExcelShape(drawings, node);
		}
		if (node.SelectSingleNode("xdr:pic", drawings.NameSpaceManager) != null)
		{
			return new ExcelPicture(drawings, node);
		}
		if (node.SelectSingleNode("xdr:graphicFrame", drawings.NameSpaceManager) != null)
		{
			return ExcelChart.GetChart(drawings, node);
		}
		return new ExcelDrawing(drawings, node, "");
	}

	internal static string GetTextAchoringText(eTextAnchoringType value)
	{
		return value switch
		{
			eTextAnchoringType.Bottom => "b", 
			eTextAnchoringType.Center => "ctr", 
			eTextAnchoringType.Distributed => "dist", 
			eTextAnchoringType.Justify => "just", 
			_ => "t", 
		};
	}

	internal static eTextAnchoringType GetTextAchoringEnum(string text)
	{
		return text switch
		{
			"b" => eTextAnchoringType.Bottom, 
			"ctr" => eTextAnchoringType.Center, 
			"dist" => eTextAnchoringType.Distributed, 
			"just" => eTextAnchoringType.Justify, 
			_ => eTextAnchoringType.Top, 
		};
	}

	internal static string GetTextVerticalText(eTextVerticalType value)
	{
		return value switch
		{
			eTextVerticalType.EastAsianVertical => "eaVert", 
			eTextVerticalType.MongolianVertical => "mongolianVert", 
			eTextVerticalType.Vertical => "vert", 
			eTextVerticalType.Vertical270 => "vert270", 
			eTextVerticalType.WordArtVertical => "wordArtVert", 
			eTextVerticalType.WordArtVerticalRightToLeft => "wordArtVertRtl", 
			_ => "horz", 
		};
	}

	internal static eTextVerticalType GetTextVerticalEnum(string text)
	{
		return text switch
		{
			"eaVert" => eTextVerticalType.EastAsianVertical, 
			"mongolianVert" => eTextVerticalType.MongolianVertical, 
			"vert" => eTextVerticalType.Vertical, 
			"vert270" => eTextVerticalType.Vertical270, 
			"wordArtVert" => eTextVerticalType.WordArtVertical, 
			"wordArtVertRtl" => eTextVerticalType.WordArtVerticalRightToLeft, 
			_ => eTextVerticalType.Horizontal, 
		};
	}

	internal int GetPixelLeft()
	{
		decimal maxFontWidth = _drawings.Worksheet.Workbook.MaxFontWidth;
		int num = 0;
		for (int i = 0; i < From.Column; i++)
		{
			num += (int)decimal.Truncate((256m * GetColumnWidth(i + 1) + decimal.Truncate(128m / maxFontWidth)) / 256m * maxFontWidth);
		}
		return num + From.ColumnOff / 9525;
	}

	internal int GetPixelTop()
	{
		_ = _drawings.Worksheet;
		int num = 0;
		for (int i = 0; i < From.Row; i++)
		{
			num += (int)(GetRowHeight(i + 1) / 0.75);
		}
		return num + From.RowOff / 9525;
	}

	internal int GetPixelWidth()
	{
		decimal maxFontWidth = _drawings.Worksheet.Workbook.MaxFontWidth;
		int num = -From.ColumnOff / 9525;
		for (int i = From.Column + 1; i <= To.Column; i++)
		{
			num += (int)decimal.Truncate((256m * GetColumnWidth(i) + decimal.Truncate(128m / maxFontWidth)) / 256m * maxFontWidth);
		}
		return num + Convert.ToInt32(Math.Round(Convert.ToDouble(To.ColumnOff) / 9525.0, 0));
	}

	internal int GetPixelHeight()
	{
		_ = _drawings.Worksheet;
		int num = -(From.RowOff / 9525);
		for (int i = From.Row + 1; i <= To.Row; i++)
		{
			num += (int)(GetRowHeight(i) / 0.75);
		}
		return num + Convert.ToInt32(Math.Round(Convert.ToDouble(To.RowOff) / 9525.0, 0));
	}

	private decimal GetColumnWidth(int col)
	{
		ExcelWorksheet worksheet = _drawings.Worksheet;
		if (!(worksheet.GetValueInner(0, col) is ExcelColumn))
		{
			return (decimal)worksheet.DefaultColWidth;
		}
		return (decimal)worksheet.Column(col).VisualWidth;
	}

	private double GetRowHeight(int row)
	{
		ExcelWorksheet worksheet = _drawings.Worksheet;
		object value = null;
		if (worksheet.ExistsValueInner(row, 0, ref value) && value != null)
		{
			RowInternal rowInternal = (RowInternal)value;
			if (rowInternal.Height >= 0.0 && rowInternal.CustomHeight)
			{
				return rowInternal.Height;
			}
			return GetRowHeightFromCellFonts(row, worksheet);
		}
		return GetRowHeightFromCellFonts(row, worksheet);
	}

	private double GetRowHeightFromCellFonts(int row, ExcelWorksheet ws)
	{
		double defaultRowHeight = ws.DefaultRowHeight;
		if (double.IsNaN(defaultRowHeight) || !ws.CustomHeight)
		{
			double num = defaultRowHeight;
			CellsStoreEnumerator<ExcelCoreValue> cellsStoreEnumerator = new CellsStoreEnumerator<ExcelCoreValue>(_drawings.Worksheet._values, row, 0, row, 16384);
			ExcelStyles styles = _drawings.Worksheet.Workbook.Styles;
			while (cellsStoreEnumerator.Next())
			{
				ExcelXfs excelXfs = styles.CellXfs[cellsStoreEnumerator.Value._styleId];
				ExcelFontXml excelFontXml = styles.Fonts[excelXfs.FontId];
				double num2 = (double)ExcelFontXml.GetFontHeight(excelFontXml.Name, excelFontXml.Size) * 0.75;
				if (num2 > num)
				{
					num = num2;
				}
			}
			return num;
		}
		return defaultRowHeight;
	}

	internal void SetPixelTop(int pixels)
	{
		_doNotAdjust = true;
		_ = _drawings.Worksheet.Workbook.MaxFontWidth;
		int num = 0;
		int i = (int)(GetRowHeight(1) / 0.75);
		int num2;
		for (num2 = 2; i < pixels; i += (int)(GetRowHeight(num2++) / 0.75))
		{
			num = i;
		}
		if (i == pixels)
		{
			From.Row = num2 - 1;
			From.RowOff = 0;
		}
		else
		{
			From.Row = num2 - 2;
			From.RowOff = (pixels - num) * 9525;
		}
		_top = pixels;
		_doNotAdjust = false;
	}

	internal void SetPixelLeft(int pixels)
	{
		_doNotAdjust = true;
		decimal maxFontWidth = _drawings.Worksheet.Workbook.MaxFontWidth;
		int num = 0;
		int i = (int)decimal.Truncate((256m * GetColumnWidth(1) + decimal.Truncate(128m / maxFontWidth)) / 256m * maxFontWidth);
		int num2;
		for (num2 = 2; i < pixels; i += (int)decimal.Truncate((256m * GetColumnWidth(num2++) + decimal.Truncate(128m / maxFontWidth)) / 256m * maxFontWidth))
		{
			num = i;
		}
		if (i == pixels)
		{
			From.Column = num2 - 1;
			From.ColumnOff = 0;
		}
		else
		{
			From.Column = num2 - 2;
			From.ColumnOff = (pixels - num) * 9525;
		}
		_left = pixels;
		_doNotAdjust = false;
	}

	internal void SetPixelHeight(int pixels)
	{
		SetPixelHeight(pixels, 96f);
	}

	internal void SetPixelHeight(int pixels, float dpi)
	{
		_doNotAdjust = true;
		_ = _drawings.Worksheet;
		pixels = (int)((double)((float)pixels / (dpi / 96f)) + 0.5);
		int num = pixels - ((int)(GetRowHeight(From.Row + 1) / 0.75) - From.RowOff / 9525);
		int num2 = pixels;
		int num3 = From.Row + 1;
		while (num >= 0)
		{
			num2 = num;
			num -= (int)(GetRowHeight(++num3) / 0.75);
		}
		To.Row = num3 - 1;
		if (From.Row == To.Row)
		{
			To.RowOff = From.RowOff + pixels * 9525;
		}
		else
		{
			To.RowOff = num2 * 9525;
		}
		_doNotAdjust = false;
	}

	internal void SetPixelWidth(int pixels)
	{
		SetPixelWidth(pixels, 96f);
	}

	internal void SetPixelWidth(int pixels, float dpi)
	{
		_doNotAdjust = true;
		decimal maxFontWidth = _drawings.Worksheet.Workbook.MaxFontWidth;
		pixels = (int)((double)((float)pixels / (dpi / 96f)) + 0.5);
		int num = pixels - ((int)decimal.Truncate((256m * GetColumnWidth(From.Column + 1) + decimal.Truncate(128m / maxFontWidth)) / 256m * maxFontWidth) - From.ColumnOff / 9525);
		int num2 = From.ColumnOff / 9525 + pixels;
		int num3 = From.Column + 2;
		while (num >= 0)
		{
			num2 = num;
			num -= (int)decimal.Truncate((256m * GetColumnWidth(num3++) + decimal.Truncate(128m / maxFontWidth)) / 256m * maxFontWidth);
		}
		To.Column = num3 - 2;
		To.ColumnOff = num2 * 9525;
		_doNotAdjust = false;
	}

	public void SetPosition(int PixelTop, int PixelLeft)
	{
		_doNotAdjust = true;
		if (_width == int.MinValue)
		{
			_width = GetPixelWidth();
			_height = GetPixelHeight();
		}
		SetPixelTop(PixelTop);
		SetPixelLeft(PixelLeft);
		SetPixelWidth(_width);
		SetPixelHeight(_height);
		_doNotAdjust = false;
	}

	public void SetPosition(int Row, int RowOffsetPixels, int Column, int ColumnOffsetPixels)
	{
		if (RowOffsetPixels < -60)
		{
			throw new ArgumentException("Minimum negative offset is -60.", "RowOffsetPixels");
		}
		if (ColumnOffsetPixels < -60)
		{
			throw new ArgumentException("Minimum negative offset is -60.", "ColumnOffsetPixels");
		}
		_doNotAdjust = true;
		if (_width == int.MinValue)
		{
			_width = GetPixelWidth();
			_height = GetPixelHeight();
		}
		From.Row = Row;
		From.RowOff = RowOffsetPixels * 9525;
		From.Column = Column;
		From.ColumnOff = ColumnOffsetPixels * 9525;
		SetPixelWidth(_width);
		SetPixelHeight(_height);
		_doNotAdjust = false;
	}

	public virtual void SetSize(int Percent)
	{
		_doNotAdjust = true;
		if (_width == int.MinValue)
		{
			_width = GetPixelWidth();
			_height = GetPixelHeight();
		}
		_width = (int)((decimal)_width * ((decimal)Percent / 100m));
		_height = (int)((decimal)_height * ((decimal)Percent / 100m));
		SetPixelWidth(_width, 96f);
		SetPixelHeight(_height, 96f);
		_doNotAdjust = false;
	}

	public void SetSize(int PixelWidth, int PixelHeight)
	{
		_doNotAdjust = true;
		_width = PixelWidth;
		_height = PixelHeight;
		SetPixelWidth(PixelWidth);
		SetPixelHeight(PixelHeight);
		_doNotAdjust = false;
	}

	internal virtual void DeleteMe()
	{
		base.TopNode.ParentNode.RemoveChild(base.TopNode);
	}

	public virtual void Dispose()
	{
		_topNode = null;
	}

	internal void GetPositionSize()
	{
		if (!_doNotAdjust)
		{
			_top = GetPixelTop();
			_left = GetPixelLeft();
			_height = GetPixelHeight();
			_width = GetPixelWidth();
		}
	}

	public void AdjustPositionAndSize()
	{
		if (_drawings.Worksheet.Workbook._package.DoAdjustDrawings)
		{
			if (EditAs == eEditAs.Absolute)
			{
				SetPixelLeft(_left);
				SetPixelTop(_top);
			}
			if (EditAs == eEditAs.Absolute || EditAs == eEditAs.OneCell)
			{
				SetPixelHeight(_height);
				SetPixelWidth(_width);
			}
		}
	}
}
