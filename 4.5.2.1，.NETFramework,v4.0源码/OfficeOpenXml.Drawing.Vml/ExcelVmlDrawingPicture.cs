using System;
using System.Drawing;
using System.Globalization;
using System.Xml;
using OfficeOpenXml.Packaging;

namespace OfficeOpenXml.Drawing.Vml;

public class ExcelVmlDrawingPicture : ExcelVmlDrawingBase
{
	private ExcelWorksheet _worksheet;

	public string Position => GetXmlNodeString("@id");

	public double Width
	{
		get
		{
			return GetStyleProp("width");
		}
		set
		{
			SetStyleProp("width", value.ToString(CultureInfo.InvariantCulture) + "pt");
		}
	}

	public double Height
	{
		get
		{
			return GetStyleProp("height");
		}
		set
		{
			SetStyleProp("height", value.ToString(CultureInfo.InvariantCulture) + "pt");
		}
	}

	public double Left
	{
		get
		{
			return GetStyleProp("left");
		}
		set
		{
			SetStyleProp("left", value.ToString(CultureInfo.InvariantCulture));
		}
	}

	public double Top
	{
		get
		{
			return GetStyleProp("top");
		}
		set
		{
			SetStyleProp("top", value.ToString(CultureInfo.InvariantCulture));
		}
	}

	public string Title
	{
		get
		{
			return GetXmlNodeString("v:imagedata/@o:title");
		}
		set
		{
			SetXmlNodeString("v:imagedata/@o:title", value);
		}
	}

	public Image Image
	{
		get
		{
			ZipPackage package = _worksheet._package.Package;
			if (package.PartExists(ImageUri))
			{
				return Image.FromStream(package.GetPart(ImageUri).GetStream());
			}
			return null;
		}
	}

	internal Uri ImageUri { get; set; }

	internal string RelId
	{
		get
		{
			return GetXmlNodeString("v:imagedata/@o:relid");
		}
		set
		{
			SetXmlNodeString("v:imagedata/@o:relid", value);
		}
	}

	public bool BiLevel
	{
		get
		{
			return GetXmlNodeString("v:imagedata/@bilevel") == "t";
		}
		set
		{
			if (value)
			{
				SetXmlNodeString("v:imagedata/@bilevel", "t");
			}
			else
			{
				DeleteNode("v:imagedata/@bilevel");
			}
		}
	}

	public bool GrayScale
	{
		get
		{
			return GetXmlNodeString("v:imagedata/@grayscale") == "t";
		}
		set
		{
			if (value)
			{
				SetXmlNodeString("v:imagedata/@grayscale", "t");
			}
			else
			{
				DeleteNode("v:imagedata/@grayscale");
			}
		}
	}

	public double Gain
	{
		get
		{
			string xmlNodeString = GetXmlNodeString("v:imagedata/@gain");
			return GetFracDT(xmlNodeString, 1.0);
		}
		set
		{
			if (value < 0.0)
			{
				throw new ArgumentOutOfRangeException("Value must be positive");
			}
			if (value == 1.0)
			{
				DeleteNode("v:imagedata/@gamma");
			}
			else
			{
				SetXmlNodeString("v:imagedata/@gain", value.ToString("#.0#", CultureInfo.InvariantCulture));
			}
		}
	}

	public double Gamma
	{
		get
		{
			string xmlNodeString = GetXmlNodeString("v:imagedata/@gamma");
			return GetFracDT(xmlNodeString, 0.0);
		}
		set
		{
			if (value == 0.0)
			{
				DeleteNode("v:imagedata/@gamma");
			}
			else
			{
				SetXmlNodeString("v:imagedata/@gamma", value.ToString("#.0#", CultureInfo.InvariantCulture));
			}
		}
	}

	public double BlackLevel
	{
		get
		{
			string xmlNodeString = GetXmlNodeString("v:imagedata/@blacklevel");
			return GetFracDT(xmlNodeString, 0.0);
		}
		set
		{
			if (value == 0.0)
			{
				DeleteNode("v:imagedata/@blacklevel");
			}
			else
			{
				SetXmlNodeString("v:imagedata/@blacklevel", value.ToString("#.0#", CultureInfo.InvariantCulture));
			}
		}
	}

	internal ExcelVmlDrawingPicture(XmlNode topNode, XmlNamespaceManager ns, ExcelWorksheet ws)
		: base(topNode, ns)
	{
		_worksheet = ws;
	}

	private double GetFracDT(string v, double def)
	{
		double result;
		if (v.EndsWith("f"))
		{
			v = v.Substring(0, v.Length - 1);
			if (double.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out result))
			{
				return result / 65535.0;
			}
			return def;
		}
		if (!double.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out result))
		{
			return def;
		}
		return result;
	}

	private void SetStyleProp(string propertyName, string value)
	{
		string xmlNodeString = GetXmlNodeString("@style");
		string text = "";
		bool flag = false;
		string[] array = xmlNodeString.Split(';');
		foreach (string text2 in array)
		{
			if (text2.Split(':')[0] == propertyName)
			{
				text = text + propertyName + ":" + value + ";";
				flag = true;
			}
			else
			{
				text = text + text2 + ";";
			}
		}
		if (!flag)
		{
			text = text + propertyName + ":" + value + ";";
		}
		SetXmlNodeString("@style", text.Substring(0, text.Length - 1));
	}

	private double GetStyleProp(string propertyName)
	{
		string[] array = GetXmlNodeString("@style").Split(';');
		for (int i = 0; i < array.Length; i++)
		{
			string[] array2 = array[i].Split(':');
			if (array2[0] == propertyName && array2.Length > 1)
			{
				if (double.TryParse(array2[1].EndsWith("pt") ? array2[1].Substring(0, array2[1].Length - 2) : array2[1], NumberStyles.Number, CultureInfo.InvariantCulture, out var result))
				{
					return result;
				}
				return 0.0;
			}
		}
		return 0.0;
	}
}
