using System;
using System.Drawing;
using System.Globalization;
using System.Xml;
using OfficeOpenXml.Style;

namespace OfficeOpenXml.Sparkline;

public class ExcelSparklineColor : XmlHelper, IColor
{
	public int Indexed
	{
		get
		{
			return GetXmlNodeInt("@indexed");
		}
		set
		{
			if (value < 0 || value > 65)
			{
				throw new ArgumentOutOfRangeException("Index out of range");
			}
			SetXmlNodeString("@indexed", value.ToString(CultureInfo.InvariantCulture));
		}
	}

	public string Rgb
	{
		get
		{
			return GetXmlNodeString("@rgb");
		}
		internal set
		{
			SetXmlNodeString("@rgb", value);
		}
	}

	public string Theme => GetXmlNodeString("@theme");

	public decimal Tint
	{
		get
		{
			return GetXmlNodeDecimal("@tint");
		}
		set
		{
			if (value > 1m || value < -1m)
			{
				throw new ArgumentOutOfRangeException("Value must be between -1 and 1");
			}
			SetXmlNodeString("@tint", value.ToString(CultureInfo.InvariantCulture));
		}
	}

	internal ExcelSparklineColor(XmlNamespaceManager ns, XmlNode node)
		: base(ns, node)
	{
	}

	public void SetColor(Color color)
	{
		Rgb = color.ToArgb().ToString("X");
	}
}
