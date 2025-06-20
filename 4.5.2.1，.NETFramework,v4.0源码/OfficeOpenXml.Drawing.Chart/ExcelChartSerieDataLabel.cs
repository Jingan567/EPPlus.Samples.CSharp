using System.Xml;
using OfficeOpenXml.Style;

namespace OfficeOpenXml.Drawing.Chart;

public sealed class ExcelChartSerieDataLabel : ExcelChartDataLabel
{
	private const string positionPath = "c:dLblPos/@val";

	private ExcelDrawingFill _fill;

	private ExcelDrawingBorder _border;

	private ExcelTextFont _font;

	public eLabelPosition Position
	{
		get
		{
			return GetPosEnum(GetXmlNodeString("c:dLblPos/@val"));
		}
		set
		{
			SetXmlNodeString("c:dLblPos/@val", GetPosText(value));
		}
	}

	public new ExcelDrawingFill Fill
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

	public new ExcelDrawingBorder Border
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

	public new ExcelTextFont Font
	{
		get
		{
			if (_font == null)
			{
				if (base.TopNode.SelectSingleNode("c:txPr", base.NameSpaceManager) == null)
				{
					CreateNode("c:txPr/a:bodyPr");
					CreateNode("c:txPr/a:lstStyle");
				}
				_font = new ExcelTextFont(base.NameSpaceManager, base.TopNode, "c:txPr/a:p/a:pPr/a:defRPr", new string[14]
				{
					"spPr", "txPr", "dLblPos", "showVal", "showCatName ", "pPr", "defRPr", "solidFill", "uFill", "latin",
					"cs", "r", "rPr", "t"
				});
			}
			return _font;
		}
	}

	internal ExcelChartSerieDataLabel(XmlNamespaceManager ns, XmlNode node)
		: base(ns, node)
	{
		CreateNode("c:dLblPos/@val");
		Position = eLabelPosition.Center;
	}
}
