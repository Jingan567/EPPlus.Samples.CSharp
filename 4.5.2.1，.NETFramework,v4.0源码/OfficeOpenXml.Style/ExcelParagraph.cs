using System.Xml;

namespace OfficeOpenXml.Style;

public sealed class ExcelParagraph : ExcelTextFont
{
	private const string TextPath = "../a:t";

	public string Text
	{
		get
		{
			return GetXmlNodeString("../a:t");
		}
		set
		{
			CreateTopNode();
			SetXmlNodeString("../a:t", value);
		}
	}

	public ExcelParagraph(XmlNamespaceManager ns, XmlNode rootNode, string path, string[] schemaNodeOrder)
		: base(ns, rootNode, path + "a:rPr", schemaNodeOrder)
	{
	}
}
