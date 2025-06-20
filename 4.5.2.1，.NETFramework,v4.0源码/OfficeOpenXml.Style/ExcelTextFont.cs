using System;
using System.Drawing;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Style;

public class ExcelTextFont : XmlHelper
{
	private string _path;

	private XmlNode _rootNode;

	private string _fontLatinPath = "a:latin/@typeface";

	private string _fontCsPath = "a:cs/@typeface";

	private string _boldPath = "@b";

	private string _underLinePath = "@u";

	private string _underLineColorPath = "a:uFill/a:solidFill/a:srgbClr/@val";

	private string _italicPath = "@i";

	private string _strikePath = "@strike";

	private string _sizePath = "@sz";

	private string _colorPath = "a:solidFill/a:srgbClr/@val";

	public string LatinFont
	{
		get
		{
			return GetXmlNodeString(_fontLatinPath);
		}
		set
		{
			CreateTopNode();
			SetXmlNodeString(_fontLatinPath, value);
		}
	}

	public string ComplexFont
	{
		get
		{
			return GetXmlNodeString(_fontCsPath);
		}
		set
		{
			CreateTopNode();
			SetXmlNodeString(_fontCsPath, value);
		}
	}

	public bool Bold
	{
		get
		{
			return GetXmlNodeBool(_boldPath);
		}
		set
		{
			CreateTopNode();
			SetXmlNodeString(_boldPath, value ? "1" : "0");
		}
	}

	public eUnderLineType UnderLine
	{
		get
		{
			return TranslateUnderline(GetXmlNodeString(_underLinePath));
		}
		set
		{
			CreateTopNode();
			SetXmlNodeString(_underLinePath, TranslateUnderlineText(value));
		}
	}

	public Color UnderLineColor
	{
		get
		{
			string xmlNodeString = GetXmlNodeString(_underLineColorPath);
			if (xmlNodeString == "")
			{
				return Color.Empty;
			}
			return Color.FromArgb(int.Parse(xmlNodeString, NumberStyles.AllowHexSpecifier));
		}
		set
		{
			CreateTopNode();
			SetXmlNodeString(_underLineColorPath, value.ToArgb().ToString("X").Substring(2, 6));
		}
	}

	public bool Italic
	{
		get
		{
			return GetXmlNodeBool(_italicPath);
		}
		set
		{
			CreateTopNode();
			SetXmlNodeString(_italicPath, value ? "1" : "0");
		}
	}

	public eStrikeType Strike
	{
		get
		{
			return TranslateStrike(GetXmlNodeString(_strikePath));
		}
		set
		{
			CreateTopNode();
			SetXmlNodeString(_strikePath, TranslateStrikeText(value));
		}
	}

	public float Size
	{
		get
		{
			return GetXmlNodeInt(_sizePath) / 100;
		}
		set
		{
			CreateTopNode();
			SetXmlNodeString(_sizePath, ((int)(value * 100f)).ToString());
		}
	}

	public Color Color
	{
		get
		{
			string xmlNodeString = GetXmlNodeString(_colorPath);
			if (xmlNodeString == "")
			{
				return Color.Empty;
			}
			return Color.FromArgb(int.Parse(xmlNodeString, NumberStyles.AllowHexSpecifier));
		}
		set
		{
			CreateTopNode();
			SetXmlNodeString(_colorPath, value.ToArgb().ToString("X").Substring(2, 6));
		}
	}

	internal ExcelTextFont(XmlNamespaceManager namespaceManager, XmlNode rootNode, string path, string[] schemaNodeOrder)
		: base(namespaceManager, rootNode)
	{
		base.SchemaNodeOrder = schemaNodeOrder;
		_rootNode = rootNode;
		if (path != "")
		{
			XmlNode xmlNode = rootNode.SelectSingleNode(path, namespaceManager);
			if (xmlNode != null)
			{
				base.TopNode = xmlNode;
			}
		}
		_path = path;
	}

	protected internal void CreateTopNode()
	{
		if (_path != "" && base.TopNode == _rootNode)
		{
			CreateNode(_path);
			base.TopNode = _rootNode.SelectSingleNode(_path, base.NameSpaceManager);
		}
	}

	private eUnderLineType TranslateUnderline(string text)
	{
		switch (text)
		{
		default:
			if (text.Length != 0)
			{
				break;
			}
			return eUnderLineType.None;
		case "sng":
			return eUnderLineType.Single;
		case "dbl":
			return eUnderLineType.Double;
		case null:
			break;
		}
		return (eUnderLineType)Enum.Parse(typeof(eUnderLineType), text);
	}

	private string TranslateUnderlineText(eUnderLineType value)
	{
		switch (value)
		{
		case eUnderLineType.Single:
			return "sng";
		case eUnderLineType.Double:
			return "dbl";
		default:
		{
			string text = value.ToString();
			return text.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + text.Substring(1, text.Length - 1);
		}
		}
	}

	private eStrikeType TranslateStrike(string text)
	{
		if (!(text == "dblStrike"))
		{
			if (text == "sngStrike")
			{
				return eStrikeType.Single;
			}
			return eStrikeType.No;
		}
		return eStrikeType.Double;
	}

	private string TranslateStrikeText(eStrikeType value)
	{
		return value switch
		{
			eStrikeType.Single => "sngStrike", 
			eStrikeType.Double => "dblStrike", 
			_ => "noStrike", 
		};
	}

	public void SetFromFont(Font Font)
	{
		LatinFont = Font.Name;
		ComplexFont = Font.Name;
		Size = Font.Size;
		if (Font.Bold)
		{
			Bold = Font.Bold;
		}
		if (Font.Italic)
		{
			Italic = Font.Italic;
		}
		if (Font.Underline)
		{
			UnderLine = eUnderLineType.Single;
		}
		if (Font.Strikeout)
		{
			Strike = eStrikeType.Single;
		}
	}
}
