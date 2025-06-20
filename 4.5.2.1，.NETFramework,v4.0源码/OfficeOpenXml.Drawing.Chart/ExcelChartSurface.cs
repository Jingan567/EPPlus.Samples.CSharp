using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart;

public class ExcelChartSurface : XmlHelper
{
	private const string THICKNESS_PATH = "c:thickness/@val";

	private ExcelDrawingFill _fill;

	private ExcelDrawingBorder _border;

	public int Thickness
	{
		get
		{
			return GetXmlNodeInt("c:thickness/@val");
		}
		set
		{
			if (value < 0 && value > 9)
			{
				throw new ArgumentOutOfRangeException("Thickness out of range. (0-9)");
			}
			SetXmlNodeString("c:thickness/@val", value.ToString());
		}
	}

	public ExcelDrawingFill Fill
	{
		get
		{
			if (_fill == null)
			{
				_fill = new ExcelDrawingFill(base.NameSpaceManager, base.TopNode, "c:spPr");
			}
			return _fill;
		}
	}

	public ExcelDrawingBorder Border
	{
		get
		{
			if (_border == null)
			{
				_border = new ExcelDrawingBorder(base.NameSpaceManager, base.TopNode, "c:spPr/a:ln");
			}
			return _border;
		}
	}

	internal ExcelChartSurface(XmlNamespaceManager ns, XmlNode node)
		: base(ns, node)
	{
		base.SchemaNodeOrder = new string[3] { "thickness", "spPr", "pictureOptions" };
	}
}
