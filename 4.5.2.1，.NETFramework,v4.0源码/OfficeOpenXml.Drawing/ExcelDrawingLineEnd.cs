using System;
using System.Xml;

namespace OfficeOpenXml.Drawing;

public sealed class ExcelDrawingLineEnd : XmlHelper
{
	private string _linePath;

	private string _headEndStylePath = "xdr:sp/xdr:spPr/a:ln/a:headEnd/@type";

	private string _tailEndStylePath = "xdr:sp/xdr:spPr/a:ln/a:tailEnd/@type";

	private string _tailEndSizeWidthPath = "xdr:sp/xdr:spPr/a:ln/a:tailEnd/@w";

	private string _tailEndSizeHeightPath = "xdr:sp/xdr:spPr/a:ln/a:tailEnd/@len";

	private string _headEndSizeWidthPath = "xdr:sp/xdr:spPr/a:ln/a:headEnd/@w";

	private string _headEndSizeHeightPath = "xdr:sp/xdr:spPr/a:ln/a:headEnd/@len";

	public eEndStyle HeadEnd
	{
		get
		{
			return TranslateEndStyle(GetXmlNodeString(_headEndStylePath));
		}
		set
		{
			CreateNode(_linePath, insertFirst: false);
			SetXmlNodeString(_headEndStylePath, TranslateEndStyleText(value));
		}
	}

	public eEndStyle TailEnd
	{
		get
		{
			return TranslateEndStyle(GetXmlNodeString(_tailEndStylePath));
		}
		set
		{
			CreateNode(_linePath, insertFirst: false);
			SetXmlNodeString(_tailEndStylePath, TranslateEndStyleText(value));
		}
	}

	public eEndSize TailEndSizeWidth
	{
		get
		{
			return TranslateEndSize(GetXmlNodeString(_tailEndSizeWidthPath));
		}
		set
		{
			CreateNode(_linePath, insertFirst: false);
			SetXmlNodeString(_tailEndSizeWidthPath, TranslateEndSizeText(value));
		}
	}

	public eEndSize TailEndSizeHeight
	{
		get
		{
			return TranslateEndSize(GetXmlNodeString(_tailEndSizeHeightPath));
		}
		set
		{
			CreateNode(_linePath, insertFirst: false);
			SetXmlNodeString(_tailEndSizeHeightPath, TranslateEndSizeText(value));
		}
	}

	public eEndSize HeadEndSizeWidth
	{
		get
		{
			return TranslateEndSize(GetXmlNodeString(_headEndSizeWidthPath));
		}
		set
		{
			CreateNode(_linePath, insertFirst: false);
			SetXmlNodeString(_headEndSizeWidthPath, TranslateEndSizeText(value));
		}
	}

	public eEndSize HeadEndSizeHeight
	{
		get
		{
			return TranslateEndSize(GetXmlNodeString(_headEndSizeHeightPath));
		}
		set
		{
			CreateNode(_linePath, insertFirst: false);
			SetXmlNodeString(_headEndSizeHeightPath, TranslateEndSizeText(value));
		}
	}

	internal ExcelDrawingLineEnd(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string linePath)
		: base(nameSpaceManager, topNode)
	{
		base.SchemaNodeOrder = new string[2] { "headEnd", "tailEnd" };
		_linePath = linePath;
	}

	private string TranslateEndStyleText(eEndStyle value)
	{
		return value.ToString().ToLower();
	}

	private eEndStyle TranslateEndStyle(string text)
	{
		switch (text)
		{
		case "none":
		case "arrow":
		case "diamond":
		case "oval":
		case "stealth":
		case "triangle":
			return (eEndStyle)Enum.Parse(typeof(eEndStyle), text, ignoreCase: true);
		default:
			throw new Exception("Invalid Endstyle");
		}
	}

	private string TranslateEndSizeText(eEndSize value)
	{
		value.ToString();
		return value switch
		{
			eEndSize.Small => "sm", 
			eEndSize.Medium => "med", 
			eEndSize.Large => "lg", 
			_ => throw new Exception("Invalid Endsize"), 
		};
	}

	private eEndSize TranslateEndSize(string text)
	{
		switch (text)
		{
		case "sm":
		case "med":
		case "lg":
			return (eEndSize)Enum.Parse(typeof(eEndSize), text, ignoreCase: true);
		default:
			throw new Exception("Invalid Endsize");
		}
	}
}
