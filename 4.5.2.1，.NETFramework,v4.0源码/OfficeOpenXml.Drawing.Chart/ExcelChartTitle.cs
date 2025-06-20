using System;
using System.Xml;
using OfficeOpenXml.Style;

namespace OfficeOpenXml.Drawing.Chart;

public class ExcelChartTitle : XmlHelper
{
	private const string titlePath = "c:tx/c:rich/a:p/a:r/a:t";

	private ExcelDrawingBorder _border;

	private ExcelDrawingFill _fill;

	private string[] paragraphNodeOrder = new string[9] { "pPr", "defRPr", "solidFill", "uFill", "latin", "cs", "r", "rPr", "t" };

	private ExcelParagraphCollection _richText;

	private const string TextVerticalPath = "xdr:sp/xdr:txBody/a:bodyPr/@vert";

	public string Text
	{
		get
		{
			return RichText.Text;
		}
		set
		{
			RichText.Text = value;
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

	public ExcelTextFont Font
	{
		get
		{
			if (_richText == null || _richText.Count == 0)
			{
				RichText.Add("");
			}
			return _richText[0];
		}
	}

	public ExcelParagraphCollection RichText
	{
		get
		{
			if (_richText == null)
			{
				_richText = new ExcelParagraphCollection(base.NameSpaceManager, base.TopNode, "c:tx/c:rich/a:p", paragraphNodeOrder);
			}
			return _richText;
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
			SetXmlNodeBool("c:overlay/@val", value);
		}
	}

	public bool AnchorCtr
	{
		get
		{
			return GetXmlNodeBool("c:tx/c:rich/a:bodyPr/@anchorCtr", blankValue: false);
		}
		set
		{
			SetXmlNodeBool("c:tx/c:rich/a:bodyPr/@anchorCtr", value, removeIf: false);
		}
	}

	public eTextAnchoringType Anchor
	{
		get
		{
			return ExcelDrawing.GetTextAchoringEnum(GetXmlNodeString("c:tx/c:rich/a:bodyPr/@anchor"));
		}
		set
		{
			SetXmlNodeString("c:tx/c:rich/a:bodyPr/@anchorCtr", ExcelDrawing.GetTextAchoringText(value));
		}
	}

	public eTextVerticalType TextVertical
	{
		get
		{
			return ExcelDrawing.GetTextVerticalEnum(GetXmlNodeString("c:tx/c:rich/a:bodyPr/@vert"));
		}
		set
		{
			SetXmlNodeString("c:tx/c:rich/a:bodyPr/@vert", ExcelDrawing.GetTextVerticalText(value));
		}
	}

	public double Rotation
	{
		get
		{
			int xmlNodeInt = GetXmlNodeInt("c:tx/c:rich/a:bodyPr/@rot");
			if (xmlNodeInt < 0)
			{
				return 360 - xmlNodeInt / 60000;
			}
			return xmlNodeInt / 60000;
		}
		set
		{
			if (value < 0.0 || value > 360.0)
			{
				throw new ArgumentOutOfRangeException("Rotation must be between 0 and 360");
			}
			SetXmlNodeString("c:tx/c:rich/a:bodyPr/@rot", ((!(value > 180.0)) ? ((int)(value * 60000.0)) : ((int)((value - 360.0) * 60000.0))).ToString());
		}
	}

	internal ExcelChartTitle(XmlNamespaceManager nameSpaceManager, XmlNode node)
		: base(nameSpaceManager, node)
	{
		XmlNode xmlNode = node.SelectSingleNode("c:title", base.NameSpaceManager);
		if (xmlNode == null)
		{
			xmlNode = node.OwnerDocument.CreateElement("c", "title", "http://schemas.openxmlformats.org/drawingml/2006/chart");
			node.InsertBefore(xmlNode, node.ChildNodes[0]);
			xmlNode.InnerXml = "<c:tx><c:rich><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr sz=\"1800\" b=\"0\" /></a:pPr><a:r><a:t /></a:r></a:p></c:rich></c:tx><c:layout /><c:overlay val=\"0\" />";
		}
		base.TopNode = xmlNode;
		base.SchemaNodeOrder = new string[5] { "tx", "bodyPr", "lstStyle", "layout", "overlay" };
	}
}
