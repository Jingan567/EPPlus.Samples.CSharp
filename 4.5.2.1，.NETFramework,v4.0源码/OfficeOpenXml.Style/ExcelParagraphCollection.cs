using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Style;

public class ExcelParagraphCollection : XmlHelper, IEnumerable<ExcelParagraph>, IEnumerable
{
	private List<ExcelParagraph> _list = new List<ExcelParagraph>();

	private string _path;

	public ExcelParagraph this[int Index] => _list[Index];

	public int Count => _list.Count;

	public string Text
	{
		get
		{
			StringBuilder stringBuilder = new StringBuilder();
			foreach (ExcelParagraph item in _list)
			{
				stringBuilder.Append(item.Text);
			}
			return stringBuilder.ToString();
		}
		set
		{
			if (Count == 0)
			{
				Add(value);
				return;
			}
			this[0].Text = value;
			_ = Count;
			for (int num = Count - 1; num > 0; num--)
			{
				RemoveAt(num);
			}
		}
	}

	internal ExcelParagraphCollection(XmlNamespaceManager ns, XmlNode topNode, string path, string[] schemaNodeOrder)
		: base(ns, topNode)
	{
		XmlNodeList xmlNodeList = topNode.SelectNodes(path + "/a:r", base.NameSpaceManager);
		base.SchemaNodeOrder = schemaNodeOrder;
		if (xmlNodeList != null)
		{
			foreach (XmlNode item in xmlNodeList)
			{
				_list.Add(new ExcelParagraph(ns, item, "", schemaNodeOrder));
			}
		}
		_path = path;
	}

	public ExcelParagraph Add(string Text)
	{
		XmlDocument xmlDocument = ((!(base.TopNode is XmlDocument)) ? base.TopNode.OwnerDocument : (base.TopNode as XmlDocument));
		XmlNode xmlNode = base.TopNode.SelectSingleNode(_path, base.NameSpaceManager);
		if (xmlNode == null)
		{
			CreateNode(_path);
			xmlNode = base.TopNode.SelectSingleNode(_path, base.NameSpaceManager);
		}
		XmlElement xmlElement = xmlDocument.CreateElement("a", "r", "http://schemas.openxmlformats.org/drawingml/2006/main");
		xmlNode.AppendChild(xmlElement);
		XmlElement newChild = xmlDocument.CreateElement("a", "rPr", "http://schemas.openxmlformats.org/drawingml/2006/main");
		xmlElement.AppendChild(newChild);
		ExcelParagraph excelParagraph = new ExcelParagraph(base.NameSpaceManager, xmlElement, "", base.SchemaNodeOrder);
		excelParagraph.ComplexFont = "Calibri";
		excelParagraph.LatinFont = "Calibri";
		excelParagraph.Size = 11f;
		excelParagraph.Text = Text;
		_list.Add(excelParagraph);
		return excelParagraph;
	}

	public void Clear()
	{
		_list.Clear();
		base.TopNode.RemoveAll();
	}

	public void RemoveAt(int Index)
	{
		XmlNode xmlNode = _list[Index].TopNode;
		while (xmlNode != null && xmlNode.Name != "a:r")
		{
			xmlNode = xmlNode.ParentNode;
		}
		xmlNode.ParentNode.RemoveChild(xmlNode);
		_list.RemoveAt(Index);
	}

	public void Remove(ExcelRichText Item)
	{
		base.TopNode.RemoveChild(Item.TopNode);
	}

	IEnumerator<ExcelParagraph> IEnumerable<ExcelParagraph>.GetEnumerator()
	{
		return _list.GetEnumerator();
	}

	IEnumerator IEnumerable.GetEnumerator()
	{
		return _list.GetEnumerator();
	}
}
