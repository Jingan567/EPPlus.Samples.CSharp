using System;
using System.Globalization;
using System.Text;
using System.Xml;
using OfficeOpenXml.Style;

namespace OfficeOpenXml.Drawing;

public sealed class ExcelShape : ExcelDrawing
{
	private const string ShapeStylePath = "xdr:sp/xdr:spPr/a:prstGeom/@prst";

	private ExcelDrawingFill _fill;

	private ExcelDrawingBorder _border;

	private ExcelDrawingLineEnd _ends;

	private string[] paragraphNodeOrder = new string[9] { "pPr", "defRPr", "solidFill", "uFill", "latin", "cs", "r", "rPr", "t" };

	private const string PARAGRAPH_PATH = "xdr:sp/xdr:txBody/a:p";

	private ExcelTextFont _font;

	private const string TextPath = "xdr:sp/xdr:txBody/a:p/a:r/a:t";

	private string lockTextPath = "xdr:sp/@fLocksText";

	private ExcelParagraphCollection _richText;

	private const string TextAnchoringPath = "xdr:sp/xdr:txBody/a:bodyPr/@anchor";

	private const string TextAnchoringCtlPath = "xdr:sp/xdr:txBody/a:bodyPr/@anchorCtr";

	private const string TEXT_ALIGN_PATH = "xdr:sp/xdr:txBody/a:p/a:pPr/@algn";

	private const string INDENT_ALIGN_PATH = "xdr:sp/xdr:txBody/a:p/a:pPr/@lvl";

	private const string TextVerticalPath = "xdr:sp/xdr:txBody/a:bodyPr/@vert";

	public eShapeStyle Style
	{
		get
		{
			string xmlNodeString = GetXmlNodeString("xdr:sp/xdr:spPr/a:prstGeom/@prst");
			try
			{
				return (eShapeStyle)Enum.Parse(typeof(eShapeStyle), xmlNodeString, ignoreCase: true);
			}
			catch
			{
				throw new Exception($"Invalid shapetype {xmlNodeString}");
			}
		}
		set
		{
			string text = value.ToString();
			text = text.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + text.Substring(1, text.Length - 1);
			SetXmlNodeString("xdr:sp/xdr:spPr/a:prstGeom/@prst", text);
		}
	}

	public ExcelDrawingFill Fill
	{
		get
		{
			if (_fill == null)
			{
				_fill = new ExcelDrawingFill(base.NameSpaceManager, base.TopNode, "xdr:sp/xdr:spPr");
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
				_border = new ExcelDrawingBorder(base.NameSpaceManager, base.TopNode, "xdr:sp/xdr:spPr/a:ln");
			}
			return _border;
		}
	}

	public ExcelDrawingLineEnd LineEnds
	{
		get
		{
			if (_ends == null)
			{
				_ends = new ExcelDrawingLineEnd(base.NameSpaceManager, base.TopNode, "xdr:sp/xdr:spPr/a:ln");
			}
			return _ends;
		}
	}

	public ExcelTextFont Font
	{
		get
		{
			if (_font == null)
			{
				if (base.TopNode.SelectSingleNode("xdr:sp/xdr:txBody/a:p", base.NameSpaceManager) == null)
				{
					Text = "";
					base.TopNode.SelectSingleNode("xdr:sp/xdr:txBody/a:p", base.NameSpaceManager);
				}
				_font = new ExcelTextFont(base.NameSpaceManager, base.TopNode, "xdr:sp/xdr:txBody/a:p/a:pPr/a:defRPr", paragraphNodeOrder);
			}
			return _font;
		}
	}

	public string Text
	{
		get
		{
			return GetXmlNodeString("xdr:sp/xdr:txBody/a:p/a:r/a:t");
		}
		set
		{
			SetXmlNodeString("xdr:sp/xdr:txBody/a:p/a:r/a:t", value);
		}
	}

	public bool LockText
	{
		get
		{
			return GetXmlNodeBool(lockTextPath, blankValue: true);
		}
		set
		{
			SetXmlNodeBool(lockTextPath, value);
		}
	}

	public ExcelParagraphCollection RichText
	{
		get
		{
			if (_richText == null)
			{
				_richText = new ExcelParagraphCollection(base.NameSpaceManager, base.TopNode, "xdr:sp/xdr:txBody/a:p", paragraphNodeOrder);
			}
			return _richText;
		}
	}

	public eTextAnchoringType TextAnchoring
	{
		get
		{
			return ExcelDrawing.GetTextAchoringEnum(GetXmlNodeString("xdr:sp/xdr:txBody/a:bodyPr/@anchor"));
		}
		set
		{
			SetXmlNodeString("xdr:sp/xdr:txBody/a:bodyPr/@anchor", ExcelDrawing.GetTextAchoringText(value));
		}
	}

	public bool TextAnchoringControl
	{
		get
		{
			return GetXmlNodeBool("xdr:sp/xdr:txBody/a:bodyPr/@anchorCtr");
		}
		set
		{
			if (value)
			{
				SetXmlNodeString("xdr:sp/xdr:txBody/a:bodyPr/@anchorCtr", "1");
			}
			else
			{
				SetXmlNodeString("xdr:sp/xdr:txBody/a:bodyPr/@anchorCtr", "0");
			}
		}
	}

	public eTextAlignment TextAlignment
	{
		get
		{
			return GetXmlNodeString("xdr:sp/xdr:txBody/a:p/a:pPr/@algn") switch
			{
				"ctr" => eTextAlignment.Center, 
				"r" => eTextAlignment.Right, 
				"dist" => eTextAlignment.Distributed, 
				"just" => eTextAlignment.Justified, 
				"justLow" => eTextAlignment.JustifiedLow, 
				"thaiDist" => eTextAlignment.ThaiDistributed, 
				_ => eTextAlignment.Left, 
			};
		}
		set
		{
			switch (value)
			{
			case eTextAlignment.Right:
				SetXmlNodeString("xdr:sp/xdr:txBody/a:p/a:pPr/@algn", "r");
				break;
			case eTextAlignment.Center:
				SetXmlNodeString("xdr:sp/xdr:txBody/a:p/a:pPr/@algn", "ctr");
				break;
			case eTextAlignment.Distributed:
				SetXmlNodeString("xdr:sp/xdr:txBody/a:p/a:pPr/@algn", "dist");
				break;
			case eTextAlignment.Justified:
				SetXmlNodeString("xdr:sp/xdr:txBody/a:p/a:pPr/@algn", "just");
				break;
			case eTextAlignment.JustifiedLow:
				SetXmlNodeString("xdr:sp/xdr:txBody/a:p/a:pPr/@algn", "justLow");
				break;
			case eTextAlignment.ThaiDistributed:
				SetXmlNodeString("xdr:sp/xdr:txBody/a:p/a:pPr/@algn", "thaiDist");
				break;
			default:
				DeleteNode("xdr:sp/xdr:txBody/a:p/a:pPr/@algn");
				break;
			}
		}
	}

	public int Indent
	{
		get
		{
			return GetXmlNodeInt("xdr:sp/xdr:txBody/a:p/a:pPr/@lvl");
		}
		set
		{
			if (value < 0 || value > 8)
			{
				throw new ArgumentOutOfRangeException("Indent level must be between 0 and 8");
			}
			SetXmlNodeString("xdr:sp/xdr:txBody/a:p/a:pPr/@lvl", value.ToString());
		}
	}

	public eTextVerticalType TextVertical
	{
		get
		{
			return ExcelDrawing.GetTextVerticalEnum(GetXmlNodeString("xdr:sp/xdr:txBody/a:bodyPr/@vert"));
		}
		set
		{
			SetXmlNodeString("xdr:sp/xdr:txBody/a:bodyPr/@vert", ExcelDrawing.GetTextVerticalText(value));
		}
	}

	internal new string Id => base.Name + Text;

	internal ExcelShape(ExcelDrawings drawings, XmlNode node)
		: base(drawings, node, "xdr:sp/xdr:nvSpPr/xdr:cNvPr/@name")
	{
		init();
	}

	internal ExcelShape(ExcelDrawings drawings, XmlNode node, eShapeStyle style)
		: base(drawings, node, "xdr:sp/xdr:nvSpPr/xdr:cNvPr/@name")
	{
		init();
		XmlElement xmlElement = node.OwnerDocument.CreateElement("xdr", "sp", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
		xmlElement.SetAttribute("macro", "");
		xmlElement.SetAttribute("textlink", "");
		node.AppendChild(xmlElement);
		xmlElement.InnerXml = ShapeStartXml();
		node.AppendChild(xmlElement.OwnerDocument.CreateElement("xdr", "clientData", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"));
	}

	private void init()
	{
		base.SchemaNodeOrder = new string[11]
		{
			"prstGeom", "ln", "pPr", "defRPr", "solidFill", "uFill", "latin", "cs", "r", "rPr",
			"t"
		};
	}

	private string ShapeStartXml()
	{
		StringBuilder stringBuilder = new StringBuilder();
		stringBuilder.AppendFormat("<xdr:nvSpPr><xdr:cNvPr id=\"{0}\" name=\"{1}\" /><xdr:cNvSpPr /></xdr:nvSpPr><xdr:spPr><a:prstGeom prst=\"rect\"><a:avLst /></a:prstGeom></xdr:spPr><xdr:style><a:lnRef idx=\"2\"><a:schemeClr val=\"accent1\"><a:shade val=\"50000\" /></a:schemeClr></a:lnRef><a:fillRef idx=\"1\"><a:schemeClr val=\"accent1\" /></a:fillRef><a:effectRef idx=\"0\"><a:schemeClr val=\"accent1\" /></a:effectRef><a:fontRef idx=\"minor\"><a:schemeClr val=\"lt1\" /></a:fontRef></xdr:style><xdr:txBody><a:bodyPr vertOverflow=\"clip\" rtlCol=\"0\" anchor=\"ctr\" /><a:lstStyle /><a:p></a:p></xdr:txBody>", _id, base.Name);
		return stringBuilder.ToString();
	}
}
