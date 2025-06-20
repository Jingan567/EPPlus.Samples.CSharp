using System;
using System.Drawing;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Style;

public class ExcelRichText : XmlHelper
{
	internal delegate void CallbackDelegate();

	private CallbackDelegate _callback;

	private const string TEXT_PATH = "d:t";

	private const string BOLD_PATH = "d:rPr/d:b";

	private const string ITALIC_PATH = "d:rPr/d:i";

	private const string STRIKE_PATH = "d:rPr/d:strike";

	private const string UNDERLINE_PATH = "d:rPr/d:u";

	private const string VERT_ALIGN_PATH = "d:rPr/d:vertAlign/@val";

	private const string SIZE_PATH = "d:rPr/d:sz/@val";

	private const string FONT_PATH = "d:rPr/d:rFont/@val";

	private const string COLOR_PATH = "d:rPr/d:color/@rgb";

	public string Text
	{
		get
		{
			return GetXmlNodeString("d:t");
		}
		set
		{
			_collection.ConvertRichtext();
			SetXmlNodeString("d:t", value, removeIfBlank: false);
			if (PreserveSpace)
			{
				(base.TopNode.SelectSingleNode("d:t", base.NameSpaceManager) as XmlElement).SetAttribute("xml:space", "preserve");
			}
			if (_callback != null)
			{
				_callback();
			}
		}
	}

	public bool PreserveSpace
	{
		get
		{
			if (base.TopNode.SelectSingleNode("d:t", base.NameSpaceManager) is XmlElement xmlElement)
			{
				return xmlElement.GetAttribute("xml:space") == "preserve";
			}
			return false;
		}
		set
		{
			_collection.ConvertRichtext();
			if (base.TopNode.SelectSingleNode("d:t", base.NameSpaceManager) is XmlElement xmlElement)
			{
				if (value)
				{
					xmlElement.SetAttribute("xml:space", "preserve");
				}
				else
				{
					xmlElement.RemoveAttribute("xml:space");
				}
			}
			if (_callback != null)
			{
				_callback();
			}
		}
	}

	public bool Bold
	{
		get
		{
			return ExistNode("d:rPr/d:b");
		}
		set
		{
			_collection.ConvertRichtext();
			if (value)
			{
				CreateNode("d:rPr/d:b");
			}
			else
			{
				DeleteNode("d:rPr/d:b");
			}
			if (_callback != null)
			{
				_callback();
			}
		}
	}

	public bool Italic
	{
		get
		{
			return ExistNode("d:rPr/d:i");
		}
		set
		{
			_collection.ConvertRichtext();
			if (value)
			{
				CreateNode("d:rPr/d:i");
			}
			else
			{
				DeleteNode("d:rPr/d:i");
			}
			if (_callback != null)
			{
				_callback();
			}
		}
	}

	public bool Strike
	{
		get
		{
			return ExistNode("d:rPr/d:strike");
		}
		set
		{
			_collection.ConvertRichtext();
			if (value)
			{
				CreateNode("d:rPr/d:strike");
			}
			else
			{
				DeleteNode("d:rPr/d:strike");
			}
			if (_callback != null)
			{
				_callback();
			}
		}
	}

	public bool UnderLine
	{
		get
		{
			return ExistNode("d:rPr/d:u");
		}
		set
		{
			_collection.ConvertRichtext();
			if (value)
			{
				CreateNode("d:rPr/d:u");
			}
			else
			{
				DeleteNode("d:rPr/d:u");
			}
			if (_callback != null)
			{
				_callback();
			}
		}
	}

	public ExcelVerticalAlignmentFont VerticalAlign
	{
		get
		{
			string xmlNodeString = GetXmlNodeString("d:rPr/d:vertAlign/@val");
			if (xmlNodeString == "")
			{
				return ExcelVerticalAlignmentFont.None;
			}
			try
			{
				return (ExcelVerticalAlignmentFont)Enum.Parse(typeof(ExcelVerticalAlignmentFont), xmlNodeString, ignoreCase: true);
			}
			catch
			{
				return ExcelVerticalAlignmentFont.None;
			}
		}
		set
		{
			_collection.ConvertRichtext();
			if (value == ExcelVerticalAlignmentFont.None)
			{
				DeleteNode("d:rPr/d:vertAlign/@val");
			}
			else
			{
				SetXmlNodeString("d:rPr/d:vertAlign/@val", value.ToString().ToLowerInvariant());
			}
			if (_callback != null)
			{
				_callback();
			}
		}
	}

	public float Size
	{
		get
		{
			return Convert.ToSingle(GetXmlNodeDecimal("d:rPr/d:sz/@val"));
		}
		set
		{
			_collection.ConvertRichtext();
			SetXmlNodeString("d:rPr/d:sz/@val", value.ToString(CultureInfo.InvariantCulture));
			if (_callback != null)
			{
				_callback();
			}
		}
	}

	public string FontName
	{
		get
		{
			return GetXmlNodeString("d:rPr/d:rFont/@val");
		}
		set
		{
			_collection.ConvertRichtext();
			SetXmlNodeString("d:rPr/d:rFont/@val", value);
			if (_callback != null)
			{
				_callback();
			}
		}
	}

	public Color Color
	{
		get
		{
			string xmlNodeString = GetXmlNodeString("d:rPr/d:color/@rgb");
			if (xmlNodeString == "")
			{
				return Color.Empty;
			}
			return Color.FromArgb(int.Parse(xmlNodeString, NumberStyles.AllowHexSpecifier));
		}
		set
		{
			_collection.ConvertRichtext();
			SetXmlNodeString("d:rPr/d:color/@rgb", value.ToArgb().ToString("X"));
			if (_callback != null)
			{
				_callback();
			}
		}
	}

	public ExcelRichTextCollection _collection { get; set; }

	internal ExcelRichText(XmlNamespaceManager ns, XmlNode topNode, ExcelRichTextCollection collection)
		: base(ns, topNode)
	{
		base.SchemaNodeOrder = new string[13]
		{
			"rPr", "t", "b", "i", "strike", "u", "vertAlign", "sz", "color", "rFont",
			"family", "scheme", "charset"
		};
		_collection = collection;
	}

	internal void SetCallback(CallbackDelegate callback)
	{
		_callback = callback;
	}
}
