using System;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Drawing;

public sealed class ExcelDrawingBorder : XmlHelper
{
	private string _linePath;

	private ExcelDrawingFill _fill;

	private string _lineStylePath = "{0}/a:prstDash/@val";

	private string _lineCapPath = "{0}/@cap";

	private string _lineWidth = "{0}/@w";

	public ExcelDrawingFill Fill
	{
		get
		{
			if (_fill == null)
			{
				_fill = new ExcelDrawingFill(base.NameSpaceManager, base.TopNode, _linePath);
			}
			return _fill;
		}
	}

	public eLineStyle LineStyle
	{
		get
		{
			return TranslateLineStyle(GetXmlNodeString(_lineStylePath));
		}
		set
		{
			CreateNode(_linePath, insertFirst: false);
			SetXmlNodeString(_lineStylePath, TranslateLineStyleText(value));
		}
	}

	public eLineCap LineCap
	{
		get
		{
			return TranslateLineCap(GetXmlNodeString(_lineCapPath));
		}
		set
		{
			CreateNode(_linePath, insertFirst: false);
			SetXmlNodeString(_lineCapPath, TranslateLineCapText(value));
		}
	}

	public int Width
	{
		get
		{
			return GetXmlNodeInt(_lineWidth) / 12700;
		}
		set
		{
			SetXmlNodeString(_lineWidth, (value * 12700).ToString());
		}
	}

	internal ExcelDrawingBorder(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string linePath)
		: base(nameSpaceManager, topNode)
	{
		base.SchemaNodeOrder = new string[19]
		{
			"chart", "tickLblPos", "spPr", "txPr", "crossAx", "printSettings", "showVal", "showCatName", "showSerName", "showPercent",
			"separator", "showLeaderLines", "noFill", "solidFill", "blipFill", "gradFill", "noFill", "pattFill", "prstDash"
		};
		_linePath = linePath;
		_lineStylePath = string.Format(_lineStylePath, linePath);
		_lineCapPath = string.Format(_lineCapPath, linePath);
		_lineWidth = string.Format(_lineWidth, linePath);
	}

	private string TranslateLineStyleText(eLineStyle value)
	{
		string text = value.ToString();
		switch (value)
		{
		case eLineStyle.Dash:
		case eLineStyle.DashDot:
		case eLineStyle.Dot:
		case eLineStyle.Solid:
			return text.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + text.Substring(1, text.Length - 1);
		case eLineStyle.LongDash:
		case eLineStyle.LongDashDot:
		case eLineStyle.LongDashDotDot:
			return "lg" + text.Substring(4, text.Length - 4);
		case eLineStyle.SystemDash:
		case eLineStyle.SystemDashDot:
		case eLineStyle.SystemDashDotDot:
		case eLineStyle.SystemDot:
			return "sys" + text.Substring(6, text.Length - 6);
		default:
			throw new Exception("Invalid Linestyle");
		}
	}

	private eLineStyle TranslateLineStyle(string text)
	{
		switch (text)
		{
		case "dash":
		case "dot":
		case "dashDot":
		case "solid":
			return (eLineStyle)Enum.Parse(typeof(eLineStyle), text, ignoreCase: true);
		case "lgDash":
		case "lgDashDot":
		case "lgDashDotDot":
			return (eLineStyle)Enum.Parse(typeof(eLineStyle), "Long" + text.Substring(2, text.Length - 2));
		case "sysDash":
		case "sysDashDot":
		case "sysDashDotDot":
		case "sysDot":
			return (eLineStyle)Enum.Parse(typeof(eLineStyle), "System" + text.Substring(3, text.Length - 3));
		default:
			throw new Exception("Invalid Linestyle");
		}
	}

	private string TranslateLineCapText(eLineCap value)
	{
		return value switch
		{
			eLineCap.Round => "rnd", 
			eLineCap.Square => "sq", 
			_ => "flat", 
		};
	}

	private eLineCap TranslateLineCap(string text)
	{
		if (!(text == "rnd"))
		{
			if (text == "sq")
			{
				return eLineCap.Square;
			}
			return eLineCap.Flat;
		}
		return eLineCap.Round;
	}
}
