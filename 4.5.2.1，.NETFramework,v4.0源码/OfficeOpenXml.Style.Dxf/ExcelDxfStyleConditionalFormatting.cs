using System;
using System.Drawing;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Style.Dxf;

public class ExcelDxfStyleConditionalFormatting : DxfStyleBase<ExcelDxfStyleConditionalFormatting>
{
	private XmlHelperInstance _helper;

	internal int DxfId { get; set; }

	public ExcelDxfFontBase Font { get; set; }

	public ExcelDxfNumberFormat NumberFormat { get; set; }

	public ExcelDxfFill Fill { get; set; }

	public ExcelDxfBorderBase Border { get; set; }

	protected internal override string Id => NumberFormat.Id + Font.Id + Border.Id + Fill.Id + (base.AllowChange ? "" : DxfId.ToString());

	protected internal override bool HasValue
	{
		get
		{
			if (!Font.HasValue && !NumberFormat.HasValue && !Fill.HasValue)
			{
				return Border.HasValue;
			}
			return true;
		}
	}

	internal ExcelDxfStyleConditionalFormatting(XmlNamespaceManager nameSpaceManager, XmlNode topNode, ExcelStyles styles)
		: base(styles)
	{
		NumberFormat = new ExcelDxfNumberFormat(_styles);
		Font = new ExcelDxfFontBase(_styles);
		Border = new ExcelDxfBorderBase(_styles);
		Fill = new ExcelDxfFill(_styles);
		if (topNode != null)
		{
			_helper = new XmlHelperInstance(nameSpaceManager, topNode);
			NumberFormat.NumFmtID = _helper.GetXmlNodeInt("d:numFmt/@numFmtId");
			NumberFormat.Format = _helper.GetXmlNodeString("d:numFmt/@formatCode");
			if (NumberFormat.NumFmtID < 164 && string.IsNullOrEmpty(NumberFormat.Format))
			{
				NumberFormat.Format = ExcelNumberFormat.GetFromBuildInFromID(NumberFormat.NumFmtID);
			}
			Font.Bold = _helper.GetXmlNodeBoolNullable("d:font/d:b/@val");
			Font.Italic = _helper.GetXmlNodeBoolNullable("d:font/d:i/@val");
			Font.Strike = _helper.GetXmlNodeBoolNullable("d:font/d:strike");
			Font.Underline = GetUnderLineEnum(_helper.GetXmlNodeString("d:font/d:u/@val"));
			Font.Color = GetColor(_helper, "d:font/d:color");
			Border.Left = GetBorderItem(_helper, "d:border/d:left");
			Border.Right = GetBorderItem(_helper, "d:border/d:right");
			Border.Bottom = GetBorderItem(_helper, "d:border/d:bottom");
			Border.Top = GetBorderItem(_helper, "d:border/d:top");
			Fill.PatternType = GetPatternTypeEnum(_helper.GetXmlNodeString("d:fill/d:patternFill/@patternType"));
			Fill.BackgroundColor = GetColor(_helper, "d:fill/d:patternFill/d:bgColor/");
			Fill.PatternColor = GetColor(_helper, "d:fill/d:patternFill/d:fgColor/");
		}
		else
		{
			_helper = new XmlHelperInstance(nameSpaceManager);
		}
		_helper.SchemaNodeOrder = new string[4] { "font", "numFmt", "fill", "border" };
	}

	private ExcelDxfBorderItem GetBorderItem(XmlHelperInstance helper, string path)
	{
		return new ExcelDxfBorderItem(_styles)
		{
			Style = GetBorderStyleEnum(helper.GetXmlNodeString(path + "/@style")),
			Color = GetColor(helper, path + "/d:color")
		};
	}

	private ExcelBorderStyle GetBorderStyleEnum(string style)
	{
		if (style == "")
		{
			return ExcelBorderStyle.None;
		}
		string value = style.Substring(0, 1).ToUpper(CultureInfo.InvariantCulture) + style.Substring(1, style.Length - 1);
		try
		{
			return (ExcelBorderStyle)Enum.Parse(typeof(ExcelBorderStyle), value);
		}
		catch
		{
			return ExcelBorderStyle.None;
		}
	}

	private ExcelFillStyle GetPatternTypeEnum(string patternType)
	{
		if (patternType == "")
		{
			return ExcelFillStyle.None;
		}
		patternType = patternType.Substring(0, 1).ToUpper(CultureInfo.InvariantCulture) + patternType.Substring(1, patternType.Length - 1);
		try
		{
			return (ExcelFillStyle)Enum.Parse(typeof(ExcelFillStyle), patternType);
		}
		catch
		{
			return ExcelFillStyle.None;
		}
	}

	private ExcelDxfColor GetColor(XmlHelperInstance helper, string path)
	{
		ExcelDxfColor excelDxfColor = new ExcelDxfColor(_styles);
		excelDxfColor.Theme = helper.GetXmlNodeIntNull(path + "/@theme");
		excelDxfColor.Index = helper.GetXmlNodeIntNull(path + "/@indexed");
		string xmlNodeString = helper.GetXmlNodeString(path + "/@rgb");
		if (xmlNodeString != "")
		{
			excelDxfColor.Color = Color.FromArgb(int.Parse(xmlNodeString.Substring(0, 2), NumberStyles.AllowHexSpecifier), int.Parse(xmlNodeString.Substring(2, 2), NumberStyles.AllowHexSpecifier), int.Parse(xmlNodeString.Substring(4, 2), NumberStyles.AllowHexSpecifier), int.Parse(xmlNodeString.Substring(6, 2), NumberStyles.AllowHexSpecifier));
		}
		excelDxfColor.Auto = helper.GetXmlNodeBoolNullable(path + "/@auto");
		excelDxfColor.Tint = helper.GetXmlNodeDoubleNull(path + "/@tint");
		return excelDxfColor;
	}

	private ExcelUnderLineType? GetUnderLineEnum(string value)
	{
		return value.ToLower(CultureInfo.InvariantCulture) switch
		{
			"single" => ExcelUnderLineType.Single, 
			"double" => ExcelUnderLineType.Double, 
			"singleaccounting" => ExcelUnderLineType.SingleAccounting, 
			"doubleaccounting" => ExcelUnderLineType.DoubleAccounting, 
			_ => null, 
		};
	}

	protected internal override ExcelDxfStyleConditionalFormatting Clone()
	{
		return new ExcelDxfStyleConditionalFormatting(_helper.NameSpaceManager, null, _styles)
		{
			Font = Font.Clone(),
			NumberFormat = NumberFormat.Clone(),
			Fill = Fill.Clone(),
			Border = Border.Clone()
		};
	}

	protected internal override void CreateNodes(XmlHelper helper, string path)
	{
		if (Font.HasValue)
		{
			Font.CreateNodes(helper, "d:font");
		}
		if (NumberFormat.HasValue)
		{
			NumberFormat.CreateNodes(helper, "d:numFmt");
		}
		if (Fill.HasValue)
		{
			Fill.CreateNodes(helper, "d:fill");
		}
		if (Border.HasValue)
		{
			Border.CreateNodes(helper, "d:border");
		}
	}
}
