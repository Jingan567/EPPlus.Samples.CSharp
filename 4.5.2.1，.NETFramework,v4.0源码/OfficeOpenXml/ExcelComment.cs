using System.Xml;
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.Style;

namespace OfficeOpenXml;

public class ExcelComment : ExcelVmlDrawingComment
{
	internal XmlHelper _commentHelper;

	private string _text;

	private const string AUTHORS_PATH = "d:comments/d:authors";

	private const string AUTHOR_PATH = "d:comments/d:authors/d:author";

	public string Author
	{
		get
		{
			int xmlNodeInt = _commentHelper.GetXmlNodeInt("@authorId");
			return _commentHelper.TopNode.OwnerDocument.SelectSingleNode(string.Format("{0}[{1}]", "d:comments/d:authors/d:author", xmlNodeInt + 1), _commentHelper.NameSpaceManager).InnerText;
		}
		set
		{
			int author = GetAuthor(value);
			_commentHelper.SetXmlNodeString("@authorId", author.ToString());
		}
	}

	public string Text
	{
		get
		{
			if (!string.IsNullOrEmpty(RichText.Text))
			{
				return RichText.Text;
			}
			return _text;
		}
		set
		{
			RichText.Text = value;
		}
	}

	public ExcelRichText Font
	{
		get
		{
			if (RichText.Count > 0)
			{
				return RichText[0];
			}
			return null;
		}
	}

	public ExcelRichTextCollection RichText { get; set; }

	internal string Reference
	{
		get
		{
			return _commentHelper.GetXmlNodeString("@ref");
		}
		set
		{
			ExcelAddressBase excelAddressBase = new ExcelAddressBase(value);
			int num = excelAddressBase._fromRow - base.Range._fromRow;
			int num2 = excelAddressBase._fromCol - base.Range._fromCol;
			base.Range.Address = value;
			_commentHelper.SetXmlNodeString("@ref", value);
			base.From.Row += num;
			base.To.Row += num;
			base.From.Column += num2;
			base.To.Column += num2;
			base.Row = base.Range._fromRow - 1;
			base.Column = base.Range._fromCol - 1;
		}
	}

	internal ExcelComment(XmlNamespaceManager ns, XmlNode commentTopNode, ExcelRangeBase cell)
		: base(null, cell, cell.Worksheet.VmlDrawingsComments.NameSpaceManager)
	{
		_commentHelper = XmlHelperFactory.Create(ns, commentTopNode);
		XmlNode xmlNode = commentTopNode.SelectSingleNode("d:text", ns);
		if (xmlNode == null)
		{
			xmlNode = commentTopNode.OwnerDocument.CreateElement("text", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
			commentTopNode.AppendChild(xmlNode);
		}
		if (!cell.Worksheet._vmlDrawings.ContainsKey(ExcelCellBase.GetCellID(cell.Worksheet.SheetID, cell.Start.Row, cell.Start.Column)))
		{
			cell.Worksheet._vmlDrawings.Add(cell);
		}
		base.TopNode = cell.Worksheet.VmlDrawingsComments[ExcelCellBase.GetCellID(cell.Worksheet.SheetID, cell.Start.Row, cell.Start.Column)].TopNode;
		RichText = new ExcelRichTextCollection(ns, xmlNode);
		XmlNode xmlNode2 = xmlNode.SelectSingleNode("d:t", ns);
		if (xmlNode2 != null)
		{
			_text = xmlNode2.InnerText;
		}
	}

	private int GetAuthor(string value)
	{
		int num = 0;
		bool flag = false;
		foreach (XmlElement item in _commentHelper.TopNode.OwnerDocument.SelectNodes("d:comments/d:authors/d:author", _commentHelper.NameSpaceManager))
		{
			if (item.InnerText == value)
			{
				flag = true;
				break;
			}
			num++;
		}
		if (!flag)
		{
			XmlElement xmlElement = _commentHelper.TopNode.OwnerDocument.CreateElement("d", "author", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
			_commentHelper.TopNode.OwnerDocument.SelectSingleNode("d:comments/d:authors", _commentHelper.NameSpaceManager).AppendChild(xmlElement);
			xmlElement.InnerText = value;
		}
		return num;
	}
}
