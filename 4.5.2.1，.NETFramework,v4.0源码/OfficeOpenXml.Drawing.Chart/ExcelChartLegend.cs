using System;
using System.Globalization;
using System.Xml;
using OfficeOpenXml.Style;

namespace OfficeOpenXml.Drawing.Chart;

public class ExcelChartLegend : XmlHelper
{
	private ExcelChart _chart;

	private const string POSITION_PATH = "c:legendPos/@val";

	private const string OVERLAY_PATH = "c:overlay/@val";

	private ExcelDrawingFill _fill;

	private ExcelDrawingBorder _border;

	private ExcelTextFont _font;

	public eLegendPosition Position
	{
		get
		{
			return GetXmlNodeString("c:legendPos/@val").ToLower(CultureInfo.InvariantCulture) switch
			{
				"t" => eLegendPosition.Top, 
				"b" => eLegendPosition.Bottom, 
				"l" => eLegendPosition.Left, 
				"tr" => eLegendPosition.TopRight, 
				_ => eLegendPosition.Right, 
			};
		}
		set
		{
			if (base.TopNode == null)
			{
				throw new Exception("Can't set position. Chart has no legend");
			}
			switch (value)
			{
			case eLegendPosition.Top:
				SetXmlNodeString("c:legendPos/@val", "t");
				break;
			case eLegendPosition.Bottom:
				SetXmlNodeString("c:legendPos/@val", "b");
				break;
			case eLegendPosition.Left:
				SetXmlNodeString("c:legendPos/@val", "l");
				break;
			case eLegendPosition.TopRight:
				SetXmlNodeString("c:legendPos/@val", "tr");
				break;
			default:
				SetXmlNodeString("c:legendPos/@val", "r");
				break;
			}
		}
	}

	public bool Overlay
	{
		get
		{
			return GetXmlNodeBool("c:overlay/@val");
		}
		set
		{
			if (base.TopNode == null)
			{
				throw new Exception("Can't set overlay. Chart has no legend");
			}
			SetXmlNodeBool("c:overlay/@val", value);
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

	public ExcelTextFont Font
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
				_font = new ExcelTextFont(base.NameSpaceManager, base.TopNode, "c:txPr/a:p/a:pPr/a:defRPr", new string[11]
				{
					"legendPos", "layout", "pPr", "defRPr", "solidFill", "uFill", "latin", "cs", "r", "rPr",
					"t"
				});
			}
			return _font;
		}
	}

	internal ExcelChartLegend(XmlNamespaceManager ns, XmlNode node, ExcelChart chart)
		: base(ns, node)
	{
		_chart = chart;
		base.SchemaNodeOrder = new string[7] { "legendPos", "layout", "overlay", "txPr", "bodyPr", "lstStyle", "spPr" };
	}

	public void Remove()
	{
		if (base.TopNode != null)
		{
			base.TopNode.ParentNode.RemoveChild(base.TopNode);
			base.TopNode = null;
		}
	}

	public void Add()
	{
		if (base.TopNode == null)
		{
			XmlHelper xmlHelper = XmlHelperFactory.Create(base.NameSpaceManager, _chart.ChartXml);
			xmlHelper.SchemaNodeOrder = _chart.SchemaNodeOrder;
			xmlHelper.CreateNode("c:chartSpace/c:chart/c:legend");
			base.TopNode = _chart.ChartXml.SelectSingleNode("c:chartSpace/c:chart/c:legend", base.NameSpaceManager);
			base.TopNode.InnerXml = "<c:legendPos val=\"r\" /><c:layout />";
		}
	}
}
