using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Vml;

public class ExcelVmlDrawingBase : XmlHelper
{
	public string Id
	{
		get
		{
			return GetXmlNodeString("@id");
		}
		set
		{
			SetXmlNodeString("@id", value);
		}
	}

	public string AlternativeText
	{
		get
		{
			return GetXmlNodeString("@alt");
		}
		set
		{
			SetXmlNodeString("@alt", value);
		}
	}

	internal ExcelVmlDrawingBase(XmlNode topNode, XmlNamespaceManager ns)
		: base(ns, topNode)
	{
		base.SchemaNodeOrder = new string[17]
		{
			"fill", "stroke", "shadow", "path", "textbox", "ClientData", "MoveWithCells", "SizeWithCells", "Anchor", "Locked",
			"AutoFill", "LockText", "TextHAlign", "TextVAlign", "Row", "Column", "Visible"
		};
	}

	protected bool GetStyle(string style, string key, out string value)
	{
		string[] array = style.Split(';');
		foreach (string text in array)
		{
			if (text.IndexOf(':') > 0)
			{
				string[] array2 = text.Split(':');
				if (array2[0] == key)
				{
					value = array2[1];
					return true;
				}
			}
			else if (text == key)
			{
				value = "";
				return true;
			}
		}
		value = "";
		return false;
	}

	protected string SetStyle(string style, string key, string value)
	{
		string[] array = style.Split(new char[1] { ';' }, StringSplitOptions.RemoveEmptyEntries);
		string text = "";
		bool flag = false;
		string[] array2 = array;
		foreach (string text2 in array2)
		{
			if (text2.Split(':')[0].Trim() == key)
			{
				if (value.Trim() != "")
				{
					text = text + key + ":" + value;
				}
				flag = true;
			}
			else
			{
				text += text2;
			}
			text += ";";
		}
		if (!flag)
		{
			return text + key + ":" + value;
		}
		return text.Substring(0, text.Length - 1);
	}
}
