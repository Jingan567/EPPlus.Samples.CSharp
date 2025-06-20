using System;
using System.Drawing;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Style.XmlAccess;

public sealed class ExcelXfs : StyleXmlHelper
{
	private ExcelStyles _styles;

	private int _xfID;

	private int _numFmtId;

	private int _fontId;

	private int _fillId;

	private int _borderId;

	private const string horizontalAlignPath = "d:alignment/@horizontal";

	private ExcelHorizontalAlignment _horizontalAlignment;

	private const string verticalAlignPath = "d:alignment/@vertical";

	private ExcelVerticalAlignment _verticalAlignment = ExcelVerticalAlignment.Bottom;

	private const string wrapTextPath = "d:alignment/@wrapText";

	private bool _wrapText;

	private string textRotationPath = "d:alignment/@textRotation";

	private int _textRotation;

	private const string lockedPath = "d:protection/@locked";

	private bool _locked = true;

	private const string hiddenPath = "d:protection/@hidden";

	private bool _hidden;

	private const string quotePrefixPath = "@quotePrefix";

	private bool _quotePrefix;

	private const string readingOrderPath = "d:alignment/@readingOrder";

	private ExcelReadingOrder _readingOrder;

	private const string shrinkToFitPath = "d:alignment/@shrinkToFit";

	private bool _shrinkToFit;

	private const string indentPath = "d:alignment/@indent";

	private int _indent;

	public int XfId
	{
		get
		{
			return _xfID;
		}
		set
		{
			_xfID = value;
		}
	}

	internal int NumberFormatId
	{
		get
		{
			return _numFmtId;
		}
		set
		{
			_numFmtId = value;
			ApplyNumberFormat = value > 0;
		}
	}

	internal int FontId
	{
		get
		{
			return _fontId;
		}
		set
		{
			_fontId = value;
		}
	}

	internal int FillId
	{
		get
		{
			return _fillId;
		}
		set
		{
			_fillId = value;
		}
	}

	internal int BorderId
	{
		get
		{
			return _borderId;
		}
		set
		{
			_borderId = value;
		}
	}

	private bool isBuildIn { get; set; }

	internal bool ApplyNumberFormat { get; set; }

	internal bool ApplyFont { get; set; }

	internal bool ApplyFill { get; set; }

	internal bool ApplyBorder { get; set; }

	internal bool ApplyAlignment { get; set; }

	internal bool ApplyProtection { get; set; }

	public ExcelStyles Styles { get; private set; }

	public ExcelNumberFormatXml Numberformat => _styles.NumberFormats[(_numFmtId >= 0) ? _numFmtId : 0];

	public ExcelFontXml Font => _styles.Fonts[(_fontId >= 0) ? _fontId : 0];

	public ExcelFillXml Fill => _styles.Fills[(_fillId >= 0) ? _fillId : 0];

	public ExcelBorderXml Border => _styles.Borders[(_borderId >= 0) ? _borderId : 0];

	public ExcelHorizontalAlignment HorizontalAlignment
	{
		get
		{
			return _horizontalAlignment;
		}
		set
		{
			_horizontalAlignment = value;
		}
	}

	public ExcelVerticalAlignment VerticalAlignment
	{
		get
		{
			return _verticalAlignment;
		}
		set
		{
			_verticalAlignment = value;
		}
	}

	public bool WrapText
	{
		get
		{
			return _wrapText;
		}
		set
		{
			_wrapText = value;
		}
	}

	public int TextRotation
	{
		get
		{
			if (_textRotation != int.MinValue)
			{
				return _textRotation;
			}
			return 0;
		}
		set
		{
			_textRotation = value;
		}
	}

	public bool Locked
	{
		get
		{
			return _locked;
		}
		set
		{
			_locked = value;
		}
	}

	public bool Hidden
	{
		get
		{
			return _hidden;
		}
		set
		{
			_hidden = value;
		}
	}

	public bool QuotePrefix
	{
		get
		{
			return _quotePrefix;
		}
		set
		{
			_quotePrefix = value;
		}
	}

	public ExcelReadingOrder ReadingOrder
	{
		get
		{
			return _readingOrder;
		}
		set
		{
			_readingOrder = value;
		}
	}

	public bool ShrinkToFit
	{
		get
		{
			return _shrinkToFit;
		}
		set
		{
			_shrinkToFit = value;
		}
	}

	public int Indent
	{
		get
		{
			if (_indent != int.MinValue)
			{
				return _indent;
			}
			return 0;
		}
		set
		{
			_indent = value;
		}
	}

	internal override string Id => XfId + "|" + NumberFormatId.ToString() + "|" + FontId.ToString() + "|" + FillId.ToString() + "|" + BorderId.ToString() + VerticalAlignment.ToString() + "|" + HorizontalAlignment.ToString() + "|" + WrapText.ToString() + "|" + ReadingOrder.ToString() + "|" + isBuildIn.ToString() + TextRotation.ToString() + Locked.ToString() + Hidden.ToString() + ShrinkToFit.ToString() + Indent.ToString() + QuotePrefix.ToString();

	internal ExcelXfs(XmlNamespaceManager nameSpaceManager, ExcelStyles styles)
		: base(nameSpaceManager)
	{
		_styles = styles;
		isBuildIn = false;
	}

	internal ExcelXfs(XmlNamespaceManager nsm, XmlNode topNode, ExcelStyles styles)
		: base(nsm, topNode)
	{
		_styles = styles;
		_xfID = GetXmlNodeInt("@xfId");
		if (_xfID == 0)
		{
			isBuildIn = true;
		}
		_numFmtId = GetXmlNodeInt("@numFmtId");
		_fontId = GetXmlNodeInt("@fontId");
		_fillId = GetXmlNodeInt("@fillId");
		_borderId = GetXmlNodeInt("@borderId");
		_readingOrder = GetReadingOrder(GetXmlNodeString("d:alignment/@readingOrder"));
		_indent = GetXmlNodeInt("d:alignment/@indent");
		_shrinkToFit = GetXmlNodeString("d:alignment/@shrinkToFit") == "1";
		_verticalAlignment = GetVerticalAlign(GetXmlNodeString("d:alignment/@vertical"));
		_horizontalAlignment = GetHorizontalAlign(GetXmlNodeString("d:alignment/@horizontal"));
		_wrapText = GetXmlNodeBool("d:alignment/@wrapText");
		_textRotation = GetXmlNodeInt(textRotationPath);
		_hidden = GetXmlNodeBool("d:protection/@hidden");
		_locked = GetXmlNodeBool("d:protection/@locked", blankValue: true);
		_quotePrefix = GetXmlNodeBool("@quotePrefix");
	}

	private ExcelReadingOrder GetReadingOrder(string value)
	{
		if (!(value == "1"))
		{
			if (value == "2")
			{
				return ExcelReadingOrder.RightToLeft;
			}
			return ExcelReadingOrder.ContextDependent;
		}
		return ExcelReadingOrder.LeftToRight;
	}

	private ExcelHorizontalAlignment GetHorizontalAlign(string align)
	{
		if (align == "")
		{
			return ExcelHorizontalAlignment.General;
		}
		align = align.Substring(0, 1).ToUpper(CultureInfo.InvariantCulture) + align.Substring(1, align.Length - 1);
		try
		{
			return (ExcelHorizontalAlignment)Enum.Parse(typeof(ExcelHorizontalAlignment), align);
		}
		catch
		{
			return ExcelHorizontalAlignment.General;
		}
	}

	private ExcelVerticalAlignment GetVerticalAlign(string align)
	{
		if (align == "")
		{
			return ExcelVerticalAlignment.Bottom;
		}
		align = align.Substring(0, 1).ToUpper(CultureInfo.InvariantCulture) + align.Substring(1, align.Length - 1);
		try
		{
			return (ExcelVerticalAlignment)Enum.Parse(typeof(ExcelVerticalAlignment), align);
		}
		catch
		{
			return ExcelVerticalAlignment.Bottom;
		}
	}

	internal void Xf_ChangedEvent(object sender, EventArgs e)
	{
	}

	internal void RegisterEvent(ExcelXfs xf)
	{
	}

	internal ExcelXfs Copy()
	{
		return Copy(_styles);
	}

	internal ExcelXfs Copy(ExcelStyles styles)
	{
		return new ExcelXfs(base.NameSpaceManager, styles)
		{
			NumberFormatId = _numFmtId,
			FontId = _fontId,
			FillId = _fillId,
			BorderId = _borderId,
			XfId = _xfID,
			ReadingOrder = _readingOrder,
			HorizontalAlignment = _horizontalAlignment,
			VerticalAlignment = _verticalAlignment,
			WrapText = _wrapText,
			ShrinkToFit = _shrinkToFit,
			Indent = _indent,
			TextRotation = _textRotation,
			Locked = _locked,
			Hidden = _hidden,
			QuotePrefix = _quotePrefix
		};
	}

	internal int GetNewID(ExcelStyleCollection<ExcelXfs> xfsCol, StyleBase styleObject, eStyleClass styleClass, eStyleProperty styleProperty, object value)
	{
		ExcelXfs excelXfs = Copy();
		switch (styleClass)
		{
		case eStyleClass.Numberformat:
			excelXfs.NumberFormatId = GetIdNumberFormat(styleProperty, value);
			styleObject.SetIndex(excelXfs.NumberFormatId);
			break;
		case eStyleClass.Font:
			excelXfs.FontId = GetIdFont(styleProperty, value);
			styleObject.SetIndex(excelXfs.FontId);
			break;
		case eStyleClass.Fill:
		case eStyleClass.FillBackgroundColor:
		case eStyleClass.FillPatternColor:
			excelXfs.FillId = GetIdFill(styleClass, styleProperty, value);
			styleObject.SetIndex(excelXfs.FillId);
			break;
		case eStyleClass.GradientFill:
		case eStyleClass.FillGradientColor1:
		case eStyleClass.FillGradientColor2:
			excelXfs.FillId = GetIdGradientFill(styleClass, styleProperty, value);
			styleObject.SetIndex(excelXfs.FillId);
			break;
		case eStyleClass.Border:
		case eStyleClass.BorderTop:
		case eStyleClass.BorderLeft:
		case eStyleClass.BorderBottom:
		case eStyleClass.BorderRight:
		case eStyleClass.BorderDiagonal:
			excelXfs.BorderId = GetIdBorder(styleClass, styleProperty, value);
			styleObject.SetIndex(excelXfs.BorderId);
			break;
		case eStyleClass.Style:
			switch (styleProperty)
			{
			case eStyleProperty.XfId:
				excelXfs.XfId = (int)value;
				break;
			case eStyleProperty.HorizontalAlign:
				excelXfs.HorizontalAlignment = (ExcelHorizontalAlignment)value;
				break;
			case eStyleProperty.VerticalAlign:
				excelXfs.VerticalAlignment = (ExcelVerticalAlignment)value;
				break;
			case eStyleProperty.WrapText:
				excelXfs.WrapText = (bool)value;
				break;
			case eStyleProperty.ReadingOrder:
				excelXfs.ReadingOrder = (ExcelReadingOrder)value;
				break;
			case eStyleProperty.ShrinkToFit:
				excelXfs.ShrinkToFit = (bool)value;
				break;
			case eStyleProperty.Indent:
				excelXfs.Indent = (int)value;
				break;
			case eStyleProperty.TextRotation:
				excelXfs.TextRotation = (int)value;
				break;
			case eStyleProperty.Locked:
				excelXfs.Locked = (bool)value;
				break;
			case eStyleProperty.Hidden:
				excelXfs.Hidden = (bool)value;
				break;
			case eStyleProperty.QuotePrefix:
				excelXfs.QuotePrefix = (bool)value;
				break;
			default:
				throw new Exception("Invalid property for class style.");
			}
			break;
		}
		int num = xfsCol.FindIndexByID(excelXfs.Id);
		if (num < 0)
		{
			return xfsCol.Add(excelXfs.Id, excelXfs);
		}
		return num;
	}

	private int GetIdBorder(eStyleClass styleClass, eStyleProperty styleProperty, object value)
	{
		ExcelBorderXml excelBorderXml = Border.Copy();
		switch (styleClass)
		{
		case eStyleClass.BorderBottom:
			SetBorderItem(excelBorderXml.Bottom, styleProperty, value);
			break;
		case eStyleClass.BorderDiagonal:
			SetBorderItem(excelBorderXml.Diagonal, styleProperty, value);
			break;
		case eStyleClass.BorderLeft:
			SetBorderItem(excelBorderXml.Left, styleProperty, value);
			break;
		case eStyleClass.BorderRight:
			SetBorderItem(excelBorderXml.Right, styleProperty, value);
			break;
		case eStyleClass.BorderTop:
			SetBorderItem(excelBorderXml.Top, styleProperty, value);
			break;
		case eStyleClass.Border:
			switch (styleProperty)
			{
			case eStyleProperty.BorderDiagonalUp:
				excelBorderXml.DiagonalUp = (bool)value;
				break;
			case eStyleProperty.BorderDiagonalDown:
				excelBorderXml.DiagonalDown = (bool)value;
				break;
			default:
				throw new Exception("Invalid property for class Border.");
			}
			break;
		default:
			throw new Exception("Invalid class/property for class Border.");
		}
		string id = excelBorderXml.Id;
		int num = _styles.Borders.FindIndexByID(id);
		if (num == int.MinValue)
		{
			return _styles.Borders.Add(id, excelBorderXml);
		}
		return num;
	}

	private void SetBorderItem(ExcelBorderItemXml excelBorderItem, eStyleProperty styleProperty, object value)
	{
		switch (styleProperty)
		{
		case eStyleProperty.Style:
			excelBorderItem.Style = (ExcelBorderStyle)value;
			break;
		case eStyleProperty.Color:
		case eStyleProperty.Tint:
		case eStyleProperty.IndexedColor:
			if (excelBorderItem.Style == ExcelBorderStyle.None)
			{
				throw new Exception("Can't set bordercolor when style is not set.");
			}
			excelBorderItem.Color.Rgb = value.ToString();
			break;
		}
	}

	private int GetIdFill(eStyleClass styleClass, eStyleProperty styleProperty, object value)
	{
		ExcelFillXml excelFillXml = Fill.Copy();
		switch (styleProperty)
		{
		case eStyleProperty.PatternType:
			if (excelFillXml is ExcelGradientFillXml)
			{
				excelFillXml = new ExcelFillXml(base.NameSpaceManager);
			}
			excelFillXml.PatternType = (ExcelFillStyle)value;
			break;
		case eStyleProperty.Color:
		case eStyleProperty.Tint:
		case eStyleProperty.IndexedColor:
		case eStyleProperty.AutoColor:
		{
			if (excelFillXml is ExcelGradientFillXml)
			{
				excelFillXml = new ExcelFillXml(base.NameSpaceManager);
			}
			if (excelFillXml.PatternType == ExcelFillStyle.None)
			{
				throw new ArgumentException("Can't set color when patterntype is not set.");
			}
			ExcelColorXml excelColorXml = ((styleClass != eStyleClass.FillPatternColor) ? excelFillXml.BackgroundColor : excelFillXml.PatternColor);
			switch (styleProperty)
			{
			case eStyleProperty.Color:
				excelColorXml.Rgb = value.ToString();
				break;
			case eStyleProperty.Tint:
				excelColorXml.Tint = (decimal)value;
				break;
			case eStyleProperty.IndexedColor:
				excelColorXml.Indexed = (int)value;
				break;
			default:
				excelColorXml.Auto = (bool)value;
				break;
			}
			break;
		}
		default:
			throw new ArgumentException("Invalid class/property for class Fill.");
		}
		string id = excelFillXml.Id;
		int num = _styles.Fills.FindIndexByID(id);
		if (num == int.MinValue)
		{
			return _styles.Fills.Add(id, excelFillXml);
		}
		return num;
	}

	private int GetIdGradientFill(eStyleClass styleClass, eStyleProperty styleProperty, object value)
	{
		ExcelGradientFillXml excelGradientFillXml;
		if (Fill is ExcelGradientFillXml)
		{
			excelGradientFillXml = (ExcelGradientFillXml)Fill.Copy();
		}
		else
		{
			excelGradientFillXml = new ExcelGradientFillXml(Fill.NameSpaceManager);
			excelGradientFillXml.GradientColor1.SetColor(Color.White);
			excelGradientFillXml.GradientColor2.SetColor(Color.FromArgb(79, 129, 189));
			excelGradientFillXml.Type = ExcelFillGradientType.Linear;
			excelGradientFillXml.Degree = 90.0;
			excelGradientFillXml.Top = double.NaN;
			excelGradientFillXml.Bottom = double.NaN;
			excelGradientFillXml.Left = double.NaN;
			excelGradientFillXml.Right = double.NaN;
		}
		switch (styleProperty)
		{
		case eStyleProperty.GradientType:
			excelGradientFillXml.Type = (ExcelFillGradientType)value;
			break;
		case eStyleProperty.GradientDegree:
			excelGradientFillXml.Degree = (double)value;
			break;
		case eStyleProperty.GradientTop:
			excelGradientFillXml.Top = (double)value;
			break;
		case eStyleProperty.GradientBottom:
			excelGradientFillXml.Bottom = (double)value;
			break;
		case eStyleProperty.GradientLeft:
			excelGradientFillXml.Left = (double)value;
			break;
		case eStyleProperty.GradientRight:
			excelGradientFillXml.Right = (double)value;
			break;
		case eStyleProperty.Color:
		case eStyleProperty.Tint:
		case eStyleProperty.IndexedColor:
		case eStyleProperty.AutoColor:
		{
			ExcelColorXml excelColorXml = ((styleClass != eStyleClass.FillGradientColor1) ? excelGradientFillXml.GradientColor2 : excelGradientFillXml.GradientColor1);
			switch (styleProperty)
			{
			case eStyleProperty.Color:
				excelColorXml.Rgb = value.ToString();
				break;
			case eStyleProperty.Tint:
				excelColorXml.Tint = (decimal)value;
				break;
			case eStyleProperty.IndexedColor:
				excelColorXml.Indexed = (int)value;
				break;
			default:
				excelColorXml.Auto = (bool)value;
				break;
			}
			break;
		}
		default:
			throw new ArgumentException("Invalid class/property for class Fill.");
		}
		string id = excelGradientFillXml.Id;
		int num = _styles.Fills.FindIndexByID(id);
		if (num == int.MinValue)
		{
			return _styles.Fills.Add(id, excelGradientFillXml);
		}
		return num;
	}

	private int GetIdNumberFormat(eStyleProperty styleProperty, object value)
	{
		if (styleProperty == eStyleProperty.Format)
		{
			ExcelNumberFormatXml obj = null;
			if (!_styles.NumberFormats.FindByID(value.ToString(), ref obj))
			{
				obj = new ExcelNumberFormatXml(base.NameSpaceManager)
				{
					Format = value.ToString(),
					NumFmtId = _styles.NumberFormats.NextId++
				};
				_styles.NumberFormats.Add(value.ToString(), obj);
			}
			return obj.NumFmtId;
		}
		throw new Exception("Invalid property for class Numberformat");
	}

	private int GetIdFont(eStyleProperty styleProperty, object value)
	{
		ExcelFontXml excelFontXml = Font.Copy();
		switch (styleProperty)
		{
		case eStyleProperty.Name:
			excelFontXml.Name = value.ToString();
			break;
		case eStyleProperty.Size:
			excelFontXml.Size = (float)value;
			break;
		case eStyleProperty.Family:
			excelFontXml.Family = (int)value;
			break;
		case eStyleProperty.Bold:
			excelFontXml.Bold = (bool)value;
			break;
		case eStyleProperty.Italic:
			excelFontXml.Italic = (bool)value;
			break;
		case eStyleProperty.Strike:
			excelFontXml.Strike = (bool)value;
			break;
		case eStyleProperty.UnderlineType:
			excelFontXml.UnderLineType = (ExcelUnderLineType)value;
			break;
		case eStyleProperty.Color:
			excelFontXml.Color.Rgb = value.ToString();
			break;
		case eStyleProperty.Tint:
			excelFontXml.Color.Tint = (decimal)value;
			break;
		case eStyleProperty.IndexedColor:
			excelFontXml.Color.Indexed = (int)value;
			break;
		case eStyleProperty.AutoColor:
			excelFontXml.Color.Auto = (bool)value;
			break;
		case eStyleProperty.VerticalAlign:
			excelFontXml.VerticalAlign = (((ExcelVerticalAlignmentFont)value == ExcelVerticalAlignmentFont.None) ? "" : value.ToString().ToLower(CultureInfo.InvariantCulture));
			break;
		default:
			throw new Exception("Invalid property for class Font");
		}
		string id = excelFontXml.Id;
		int num = _styles.Fonts.FindIndexByID(id);
		if (num == int.MinValue)
		{
			return _styles.Fonts.Add(id, excelFontXml);
		}
		return num;
	}

	internal override XmlNode CreateXmlNode(XmlNode topNode)
	{
		return CreateXmlNode(topNode, isCellStyleXsf: false);
	}

	internal XmlNode CreateXmlNode(XmlNode topNode, bool isCellStyleXsf)
	{
		base.TopNode = topNode;
		bool flag = !isCellStyleXsf && _xfID > int.MinValue && _styles.CellStyleXfs.Count > 0 && _styles.CellStyleXfs[_xfID].newID > int.MinValue;
		if (_numFmtId >= 0)
		{
			SetXmlNodeString("@numFmtId", _numFmtId.ToString());
			if (flag)
			{
				SetXmlNodeString("@applyNumberFormat", "1");
			}
		}
		if (_fontId >= 0)
		{
			SetXmlNodeString("@fontId", _styles.Fonts[_fontId].newID.ToString());
			if (flag)
			{
				SetXmlNodeString("@applyFont", "1");
			}
		}
		if (_fillId >= 0)
		{
			SetXmlNodeString("@fillId", _styles.Fills[_fillId].newID.ToString());
			if (flag)
			{
				SetXmlNodeString("@applyFill", "1");
			}
		}
		if (_borderId >= 0)
		{
			SetXmlNodeString("@borderId", _styles.Borders[_borderId].newID.ToString());
			if (flag)
			{
				SetXmlNodeString("@applyBorder", "1");
			}
		}
		if (_horizontalAlignment != 0)
		{
			SetXmlNodeString("d:alignment/@horizontal", SetAlignString(_horizontalAlignment));
		}
		if (flag)
		{
			SetXmlNodeString("@xfId", _styles.CellStyleXfs[_xfID].newID.ToString());
		}
		if (_verticalAlignment != ExcelVerticalAlignment.Bottom)
		{
			SetXmlNodeString("d:alignment/@vertical", SetAlignString(_verticalAlignment));
		}
		if (_wrapText)
		{
			SetXmlNodeString("d:alignment/@wrapText", "1");
		}
		if (_readingOrder != 0)
		{
			int readingOrder = (int)_readingOrder;
			SetXmlNodeString("d:alignment/@readingOrder", readingOrder.ToString());
		}
		if (_shrinkToFit)
		{
			SetXmlNodeString("d:alignment/@shrinkToFit", "1");
		}
		if (_indent > 0)
		{
			SetXmlNodeString("d:alignment/@indent", _indent.ToString());
		}
		if (_textRotation > 0)
		{
			SetXmlNodeString(textRotationPath, _textRotation.ToString());
		}
		if (!_locked)
		{
			SetXmlNodeString("d:protection/@locked", "0");
		}
		if (_hidden)
		{
			SetXmlNodeString("d:protection/@hidden", "1");
		}
		if (_quotePrefix)
		{
			SetXmlNodeString("@quotePrefix", "1");
		}
		return base.TopNode;
	}

	private string SetAlignString(Enum align)
	{
		string name = Enum.GetName(align.GetType(), align);
		return name.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + name.Substring(1, name.Length - 1);
	}
}
