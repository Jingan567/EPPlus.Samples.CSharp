using System.Xml;
using OfficeOpenXml.Style;

namespace OfficeOpenXml.Drawing.Chart;

public class ExcelChartDataLabel : XmlHelper
{
	private const string showValPath = "c:showVal/@val";

	private const string showCatPath = "c:showCatName/@val";

	private const string showSerPath = "c:showSerName/@val";

	private const string showPerentPath = "c:showPercent/@val";

	private const string showLeaderLinesPath = "c:showLeaderLines/@val";

	private const string showBubbleSizePath = "c:showBubbleSize/@val";

	private const string showLegendKeyPath = "c:showLegendKey/@val";

	private const string separatorPath = "c:separator";

	private ExcelDrawingFill _fill;

	private ExcelDrawingBorder _border;

	private string[] _paragraphSchemaOrder = new string[18]
	{
		"spPr", "txPr", "dLblPos", "showVal", "showCatName", "showSerName", "showPercent", "separator", "showLeaderLines", "pPr",
		"defRPr", "solidFill", "uFill", "latin", "cs", "r", "rPr", "t"
	};

	private ExcelTextFont _font;

	public bool ShowValue
	{
		get
		{
			return GetXmlNodeBool("c:showVal/@val");
		}
		set
		{
			SetXmlNodeString("c:showVal/@val", value ? "1" : "0");
		}
	}

	public bool ShowCategory
	{
		get
		{
			return GetXmlNodeBool("c:showCatName/@val");
		}
		set
		{
			SetXmlNodeString("c:showCatName/@val", value ? "1" : "0");
		}
	}

	public bool ShowSeriesName
	{
		get
		{
			return GetXmlNodeBool("c:showSerName/@val");
		}
		set
		{
			SetXmlNodeString("c:showSerName/@val", value ? "1" : "0");
		}
	}

	public bool ShowPercent
	{
		get
		{
			return GetXmlNodeBool("c:showPercent/@val");
		}
		set
		{
			SetXmlNodeString("c:showPercent/@val", value ? "1" : "0");
		}
	}

	public bool ShowLeaderLines
	{
		get
		{
			return GetXmlNodeBool("c:showLeaderLines/@val");
		}
		set
		{
			SetXmlNodeString("c:showLeaderLines/@val", value ? "1" : "0");
		}
	}

	public bool ShowBubbleSize
	{
		get
		{
			return GetXmlNodeBool("c:showBubbleSize/@val");
		}
		set
		{
			SetXmlNodeString("c:showBubbleSize/@val", value ? "1" : "0");
		}
	}

	public bool ShowLegendKey
	{
		get
		{
			return GetXmlNodeBool("c:showLegendKey/@val");
		}
		set
		{
			SetXmlNodeString("c:showLegendKey/@val", value ? "1" : "0");
		}
	}

	public string Separator
	{
		get
		{
			return GetXmlNodeString("c:separator");
		}
		set
		{
			if (string.IsNullOrEmpty(value))
			{
				DeleteNode("c:separator");
			}
			else
			{
				SetXmlNodeString("c:separator", value);
			}
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
				_font = new ExcelTextFont(base.NameSpaceManager, base.TopNode, "c:txPr/a:p/a:pPr/a:defRPr", _paragraphSchemaOrder);
			}
			return _font;
		}
	}

	internal ExcelChartDataLabel(XmlNamespaceManager ns, XmlNode node)
		: base(ns, node)
	{
		XmlNode xmlNode = node.SelectSingleNode("c:dLbls", base.NameSpaceManager);
		if (xmlNode == null)
		{
			xmlNode = node.OwnerDocument.CreateElement("c", "dLbls", "http://schemas.openxmlformats.org/drawingml/2006/chart");
			InserAfter(node, "c:marker,c:tx,c:order,c:ser", xmlNode);
			base.SchemaNodeOrder = new string[11]
			{
				"spPr", "txPr", "dLblPos", "showLegendKey", "showVal", "showCatName", "showSerName", "showPercent", "showBubbleSize", "separator",
				"showLeaderLines"
			};
			xmlNode.InnerXml = "<c:showLegendKey val=\"0\" /><c:showVal val=\"0\" /><c:showCatName val=\"0\" /><c:showSerName val=\"0\" /><c:showPercent val=\"0\" /><c:showBubbleSize val=\"0\" /> <c:separator>\r\n</c:separator><c:showLeaderLines val=\"0\" />";
		}
		base.TopNode = xmlNode;
	}

	protected string GetPosText(eLabelPosition pos)
	{
		return pos switch
		{
			eLabelPosition.Bottom => "b", 
			eLabelPosition.Center => "ctr", 
			eLabelPosition.InBase => "inBase", 
			eLabelPosition.InEnd => "inEnd", 
			eLabelPosition.Left => "l", 
			eLabelPosition.Right => "r", 
			eLabelPosition.Top => "t", 
			eLabelPosition.OutEnd => "outEnd", 
			_ => "bestFit", 
		};
	}

	protected eLabelPosition GetPosEnum(string pos)
	{
		return pos switch
		{
			"b" => eLabelPosition.Bottom, 
			"ctr" => eLabelPosition.Center, 
			"inBase" => eLabelPosition.InBase, 
			"inEnd" => eLabelPosition.InEnd, 
			"l" => eLabelPosition.Left, 
			"r" => eLabelPosition.Right, 
			"t" => eLabelPosition.Top, 
			"outEnd" => eLabelPosition.OutEnd, 
			_ => eLabelPosition.BestFit, 
		};
	}
}
