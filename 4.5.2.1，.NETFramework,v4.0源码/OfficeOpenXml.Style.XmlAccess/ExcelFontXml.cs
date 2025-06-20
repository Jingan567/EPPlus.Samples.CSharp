using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Style.XmlAccess;

public sealed class ExcelFontXml : StyleXmlHelper
{
	private const string namePath = "d:name/@val";

	private string _name;

	private const string sizePath = "d:sz/@val";

	private float _size;

	private const string familyPath = "d:family/@val";

	private int _family;

	private ExcelColorXml _color;

	private const string _colorPath = "d:color";

	private const string schemePath = "d:scheme/@val";

	private string _scheme = "";

	private const string boldPath = "d:b";

	private bool _bold;

	private const string italicPath = "d:i";

	private bool _italic;

	private const string strikePath = "d:strike";

	private bool _strike;

	private const string underLinedPath = "d:u";

	private ExcelUnderLineType _underlineType;

	private const string verticalAlignPath = "d:vertAlign/@val";

	private string _verticalAlign;

	internal override string Id => Name + "|" + Size + "|" + Family + "|" + Color.Id + "|" + Scheme + "|" + Bold.ToString() + "|" + Italic.ToString() + "|" + Strike.ToString() + "|" + VerticalAlign + "|" + UnderLineType.ToString();

	public string Name
	{
		get
		{
			return _name;
		}
		set
		{
			Scheme = "";
			_name = value;
		}
	}

	public float Size
	{
		get
		{
			return _size;
		}
		set
		{
			_size = value;
		}
	}

	public int Family
	{
		get
		{
			if (_family != int.MinValue)
			{
				return _family;
			}
			return 0;
		}
		set
		{
			_family = value;
		}
	}

	public ExcelColorXml Color
	{
		get
		{
			return _color;
		}
		internal set
		{
			_color = value;
		}
	}

	public string Scheme
	{
		get
		{
			return _scheme;
		}
		private set
		{
			_scheme = value;
		}
	}

	public bool Bold
	{
		get
		{
			return _bold;
		}
		set
		{
			_bold = value;
		}
	}

	public bool Italic
	{
		get
		{
			return _italic;
		}
		set
		{
			_italic = value;
		}
	}

	public bool Strike
	{
		get
		{
			return _strike;
		}
		set
		{
			_strike = value;
		}
	}

	public bool UnderLine
	{
		get
		{
			return UnderLineType != ExcelUnderLineType.None;
		}
		set
		{
			_underlineType = (value ? ExcelUnderLineType.Single : ExcelUnderLineType.None);
		}
	}

	public ExcelUnderLineType UnderLineType
	{
		get
		{
			return _underlineType;
		}
		set
		{
			_underlineType = value;
		}
	}

	public string VerticalAlign
	{
		get
		{
			return _verticalAlign;
		}
		set
		{
			_verticalAlign = value;
		}
	}

	internal ExcelFontXml(XmlNamespaceManager nameSpaceManager)
		: base(nameSpaceManager)
	{
		_name = "";
		_size = 0f;
		_family = int.MinValue;
		_scheme = "";
		_color = (_color = new ExcelColorXml(base.NameSpaceManager));
		_bold = false;
		_italic = false;
		_strike = false;
		_underlineType = ExcelUnderLineType.None;
		_verticalAlign = "";
	}

	internal ExcelFontXml(XmlNamespaceManager nsm, XmlNode topNode)
		: base(nsm, topNode)
	{
		_name = GetXmlNodeString("d:name/@val");
		_size = (float)GetXmlNodeDecimal("d:sz/@val");
		_family = GetXmlNodeIntNull("d:family/@val") ?? int.MinValue;
		_scheme = GetXmlNodeString("d:scheme/@val");
		_color = new ExcelColorXml(nsm, topNode.SelectSingleNode("d:color", nsm));
		_bold = GetBoolValue(topNode, "d:b");
		_italic = GetBoolValue(topNode, "d:i");
		_strike = GetBoolValue(topNode, "d:strike");
		_verticalAlign = GetXmlNodeString("d:vertAlign/@val");
		if (topNode.SelectSingleNode("d:u", base.NameSpaceManager) != null)
		{
			string xmlNodeString = GetXmlNodeString("d:u/@val");
			if (xmlNodeString == "")
			{
				_underlineType = ExcelUnderLineType.Single;
			}
			else
			{
				_underlineType = (ExcelUnderLineType)Enum.Parse(typeof(ExcelUnderLineType), xmlNodeString, ignoreCase: true);
			}
		}
		else
		{
			_underlineType = ExcelUnderLineType.None;
		}
	}

	public void SetFromFont(Font Font)
	{
		Name = Font.Name;
		Size = (int)Font.Size;
		Strike = Font.Strikeout;
		Bold = Font.Bold;
		UnderLine = Font.Underline;
		Italic = Font.Italic;
	}

	public static float GetFontHeight(string name, float size)
	{
		name = (name.StartsWith("@") ? name.Substring(1) : name);
		if (FontSize.FontHeights.ContainsKey(name))
		{
			return GetHeightByName(name, size);
		}
		return GetHeightByName("Calibri", size);
	}

	private static float GetHeightByName(string name, float size)
	{
		if (FontSize.FontHeights[name].ContainsKey(size))
		{
			return FontSize.FontHeights[name][size].Height;
		}
		float num = -1f;
		float num2 = 500f;
		foreach (KeyValuePair<float, FontSizeInfo> item in FontSize.FontHeights[name])
		{
			if (num < item.Key && item.Key < size)
			{
				num = item.Key;
			}
			if (num2 > item.Key && item.Key > size)
			{
				num2 = item.Key;
			}
		}
		if (num == num2)
		{
			return Convert.ToSingle(FontSize.FontHeights[name][num].Height);
		}
		return Convert.ToSingle(FontSize.FontHeights[name][num].Height + (FontSize.FontHeights[name][num2].Height - FontSize.FontHeights[name][num].Height) * ((size - num) / (num2 - num)));
	}

	internal ExcelFontXml Copy()
	{
		return new ExcelFontXml(base.NameSpaceManager)
		{
			Name = _name,
			Size = _size,
			Family = _family,
			Scheme = _scheme,
			Bold = _bold,
			Italic = _italic,
			UnderLineType = _underlineType,
			Strike = _strike,
			VerticalAlign = _verticalAlign,
			Color = Color.Copy()
		};
	}

	internal override XmlNode CreateXmlNode(XmlNode topElement)
	{
		base.TopNode = topElement;
		if (_bold)
		{
			CreateNode("d:b");
		}
		else
		{
			DeleteAllNode("d:b");
		}
		if (_italic)
		{
			CreateNode("d:i");
		}
		else
		{
			DeleteAllNode("d:i");
		}
		if (_strike)
		{
			CreateNode("d:strike");
		}
		else
		{
			DeleteAllNode("d:strike");
		}
		if (_underlineType == ExcelUnderLineType.None)
		{
			DeleteAllNode("d:u");
		}
		else if (_underlineType == ExcelUnderLineType.Single)
		{
			CreateNode("d:u");
		}
		else
		{
			string text = _underlineType.ToString();
			SetXmlNodeString("d:u/@val", text.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + text.Substring(1));
		}
		if (_verticalAlign != "")
		{
			SetXmlNodeString("d:vertAlign/@val", _verticalAlign.ToString());
		}
		if (_size > 0f)
		{
			SetXmlNodeString("d:sz/@val", _size.ToString(CultureInfo.InvariantCulture));
		}
		if (_color.Exists)
		{
			CreateNode("d:color");
			base.TopNode.AppendChild(_color.CreateXmlNode(base.TopNode.SelectSingleNode("d:color", base.NameSpaceManager)));
		}
		if (!string.IsNullOrEmpty(_name))
		{
			SetXmlNodeString("d:name/@val", _name);
		}
		if (_family > int.MinValue)
		{
			SetXmlNodeString("d:family/@val", _family.ToString());
		}
		if (_scheme != "")
		{
			SetXmlNodeString("d:scheme/@val", _scheme.ToString());
		}
		return base.TopNode;
	}
}
