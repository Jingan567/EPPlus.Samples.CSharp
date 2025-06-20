using System;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Style.XmlAccess;

public sealed class ExcelBorderItemXml : StyleXmlHelper
{
	private ExcelBorderStyle _borderStyle;

	private ExcelColorXml _color;

	private const string _colorPath = "d:color";

	public ExcelBorderStyle Style
	{
		get
		{
			return _borderStyle;
		}
		set
		{
			_borderStyle = value;
			Exists = true;
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

	internal override string Id
	{
		get
		{
			if (Exists)
			{
				return string.Concat(Style, Color.Id);
			}
			return "None";
		}
	}

	public bool Exists { get; private set; }

	internal ExcelBorderItemXml(XmlNamespaceManager nameSpaceManager)
		: base(nameSpaceManager)
	{
		_borderStyle = ExcelBorderStyle.None;
		_color = new ExcelColorXml(base.NameSpaceManager);
	}

	internal ExcelBorderItemXml(XmlNamespaceManager nsm, XmlNode topNode)
		: base(nsm, topNode)
	{
		if (topNode != null)
		{
			_borderStyle = GetBorderStyle(GetXmlNodeString("@style"));
			_color = new ExcelColorXml(nsm, topNode.SelectSingleNode("d:color", nsm));
			Exists = true;
		}
		else
		{
			Exists = false;
		}
	}

	private ExcelBorderStyle GetBorderStyle(string style)
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

	internal ExcelBorderItemXml Copy()
	{
		return new ExcelBorderItemXml(base.NameSpaceManager)
		{
			Style = _borderStyle,
			Color = _color.Copy()
		};
	}

	internal override XmlNode CreateXmlNode(XmlNode topNode)
	{
		base.TopNode = topNode;
		if (Style != 0)
		{
			SetXmlNodeString("@style", SetBorderString(Style));
			if (Color.Exists)
			{
				CreateNode("d:color");
				topNode.AppendChild(Color.CreateXmlNode(base.TopNode.SelectSingleNode("d:color", base.NameSpaceManager)));
			}
		}
		return base.TopNode;
	}

	private string SetBorderString(ExcelBorderStyle Style)
	{
		string name = Enum.GetName(typeof(ExcelBorderStyle), Style);
		return name.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + name.Substring(1, name.Length - 1);
	}
}
