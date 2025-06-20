using System.Xml;
using OfficeOpenXml.Style;

namespace OfficeOpenXml.Drawing.Chart;

public class ExcelChartDataTable : XmlHelper
{
	private const string showHorzBorderPath = "c:showHorzBorder/@val";

	private const string showVertBorderPath = "c:showVertBorder/@val";

	private const string showOutlinePath = "c:showOutline/@val";

	private const string showKeysPath = "c:showKeys/@val";

	private ExcelDrawingFill _fill;

	private ExcelDrawingBorder _border;

	private string[] _paragraphSchemaOrder = new string[18]
	{
		"spPr", "txPr", "dLblPos", "showVal", "showCatName", "showSerName", "showPercent", "separator", "showLeaderLines", "pPr",
		"defRPr", "solidFill", "uFill", "latin", "cs", "r", "rPr", "t"
	};

	private ExcelTextFont _font;

	public bool ShowHorizontalBorder
	{
		get
		{
			return GetXmlNodeBool("c:showHorzBorder/@val");
		}
		set
		{
			SetXmlNodeString("c:showHorzBorder/@val", value ? "1" : "0");
		}
	}

	public bool ShowVerticalBorder
	{
		get
		{
			return GetXmlNodeBool("c:showVertBorder/@val");
		}
		set
		{
			SetXmlNodeString("c:showVertBorder/@val", value ? "1" : "0");
		}
	}

	public bool ShowOutline
	{
		get
		{
			return GetXmlNodeBool("c:showOutline/@val");
		}
		set
		{
			SetXmlNodeString("c:showOutline/@val", value ? "1" : "0");
		}
	}

	public bool ShowKeys
	{
		get
		{
			return GetXmlNodeBool("c:showKeys/@val");
		}
		set
		{
			SetXmlNodeString("c:showKeys/@val", value ? "1" : "0");
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

	internal ExcelChartDataTable(XmlNamespaceManager ns, XmlNode node)
		: base(ns, node)
	{
		XmlNode xmlNode = node.SelectSingleNode("c:dTable", base.NameSpaceManager);
		if (xmlNode == null)
		{
			xmlNode = node.OwnerDocument.CreateElement("c", "dTable", "http://schemas.openxmlformats.org/drawingml/2006/chart");
			InserAfter(node, "c:valAx,c:catAx", xmlNode);
			base.SchemaNodeOrder = new string[7] { "dTable", "showHorzBorder", "showVertBorder", "showOutline", "showKeys", "spPr", "txPr" };
			xmlNode.InnerXml = "<c:showHorzBorder val=\"1\"/><c:showVertBorder val=\"1\"/><c:showOutline val=\"1\"/><c:showKeys val=\"1\"/>";
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
