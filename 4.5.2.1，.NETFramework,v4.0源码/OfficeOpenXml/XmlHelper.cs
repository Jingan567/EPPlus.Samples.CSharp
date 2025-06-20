using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Style;

namespace OfficeOpenXml;

public abstract class XmlHelper
{
	internal delegate int ChangedEventHandler(StyleBase sender, StyleChangeEventArgs e);

	internal enum eNodeInsertOrder
	{
		First,
		Last,
		After,
		Before,
		SchemaOrder
	}

	private string[] _schemaNodeOrder;

	internal XmlNamespaceManager NameSpaceManager { get; set; }

	internal XmlNode TopNode { get; set; }

	internal string[] SchemaNodeOrder
	{
		get
		{
			return _schemaNodeOrder;
		}
		set
		{
			_schemaNodeOrder = value;
		}
	}

	internal XmlHelper(XmlNamespaceManager nameSpaceManager)
	{
		TopNode = null;
		NameSpaceManager = nameSpaceManager;
	}

	internal XmlHelper(XmlNamespaceManager nameSpaceManager, XmlNode topNode)
	{
		TopNode = topNode;
		NameSpaceManager = nameSpaceManager;
	}

	internal XmlNode CreateNode(string path)
	{
		if (path == "")
		{
			return TopNode;
		}
		return CreateNode(path, insertFirst: false);
	}

	internal XmlNode CreateNode(string path, bool insertFirst, bool addNew = false)
	{
		XmlNode xmlNode = TopNode;
		XmlNode xmlNode2 = null;
		if (path.StartsWith("/"))
		{
			path = path.Substring(1);
		}
		string[] array = path.Split('/');
		for (int i = 0; i < array.Length; i++)
		{
			string text = array[i];
			XmlNode xmlNode3 = xmlNode.SelectSingleNode(text, NameSpaceManager);
			if (xmlNode3 == null || (i == text.Length - 1 && addNew))
			{
				string text2 = "";
				string[] array2 = text.Split(':');
				if (SchemaNodeOrder != null && text[0] != '@')
				{
					insertFirst = false;
					xmlNode2 = GetPrependNode(text, xmlNode);
				}
				string text3;
				string text4;
				if (array2.Length > 1)
				{
					text3 = array2[0];
					if (text3[0] == '@')
					{
						text3 = text3.Substring(1, text3.Length - 1);
					}
					text2 = NameSpaceManager.LookupNamespace(text3);
					text4 = array2[1];
				}
				else
				{
					text3 = "";
					text2 = "";
					text4 = array2[0];
				}
				if (text.StartsWith("@"))
				{
					XmlAttribute node = xmlNode.OwnerDocument.CreateAttribute(text.Substring(1, text.Length - 1), text2);
					xmlNode.Attributes.Append(node);
				}
				else
				{
					xmlNode3 = ((text3 == "") ? xmlNode.OwnerDocument.CreateElement(text4, text2) : ((!(text3 == "") && (xmlNode.OwnerDocument == null || xmlNode.OwnerDocument.DocumentElement == null || !(xmlNode.OwnerDocument.DocumentElement.NamespaceURI == text2) || !(xmlNode.OwnerDocument.DocumentElement.Prefix == ""))) ? xmlNode.OwnerDocument.CreateElement(text3, text4, text2) : xmlNode.OwnerDocument.CreateElement(text4, text2)));
					if (xmlNode2 != null)
					{
						xmlNode.InsertBefore(xmlNode3, xmlNode2);
						xmlNode2 = null;
					}
					else if (insertFirst)
					{
						xmlNode.PrependChild(xmlNode3);
					}
					else
					{
						xmlNode.AppendChild(xmlNode3);
					}
				}
			}
			xmlNode = xmlNode3;
		}
		return xmlNode;
	}

	internal XmlNode CreateComplexNode(string path)
	{
		return CreateComplexNode(TopNode, path, eNodeInsertOrder.SchemaOrder, null);
	}

	internal XmlNode CreateComplexNode(XmlNode topNode, string path)
	{
		return CreateComplexNode(topNode, path, eNodeInsertOrder.SchemaOrder, null);
	}

	internal XmlNode CreateComplexNode(XmlNode topNode, string path, eNodeInsertOrder nodeInsertOrder, XmlNode referenceNode)
	{
		if (path == null || path == string.Empty)
		{
			return topNode;
		}
		XmlNode xmlNode = topNode;
		string empty = string.Empty;
		string[] array = path.Split('/');
		foreach (string text in array)
		{
			if (text.Length <= 0)
			{
				continue;
			}
			if (text.StartsWith("@"))
			{
				string[] array2 = text.Split('=');
				string name = array2[0].Substring(1, array2[0].Length - 1);
				string text2 = null;
				if (array2.Length > 1)
				{
					text2 = array2[1].Replace("'", "").Replace("\"", "");
				}
				XmlAttribute xmlAttribute = (XmlAttribute)xmlNode.Attributes.GetNamedItem(name);
				if (text2 == string.Empty)
				{
					if (xmlAttribute != null)
					{
						xmlNode.Attributes.Remove(xmlAttribute);
					}
					continue;
				}
				if (xmlAttribute == null)
				{
					xmlAttribute = xmlNode.OwnerDocument.CreateAttribute(name);
					xmlNode.Attributes.Append(xmlAttribute);
				}
				if (text2 != null)
				{
					xmlNode.Attributes[name].Value = text2;
				}
				continue;
			}
			XmlNode xmlNode2 = xmlNode.SelectSingleNode(text, NameSpaceManager);
			if (xmlNode2 == null)
			{
				string[] array3 = text.Split(':');
				empty = string.Empty;
				string text3;
				string text4;
				if (array3.Length > 1)
				{
					text3 = array3[0];
					empty = NameSpaceManager.LookupNamespace(text3);
					text4 = array3[1];
				}
				else
				{
					text3 = string.Empty;
					empty = string.Empty;
					text4 = array3[0];
				}
				if (text4.IndexOf("[") > 0)
				{
					text4 = text4.Substring(0, text4.IndexOf("["));
				}
				xmlNode2 = ((text3 == string.Empty) ? xmlNode.OwnerDocument.CreateElement(text4, empty) : ((xmlNode.OwnerDocument == null || xmlNode.OwnerDocument.DocumentElement == null || !(xmlNode.OwnerDocument.DocumentElement.NamespaceURI == empty) || !(xmlNode.OwnerDocument.DocumentElement.Prefix == string.Empty)) ? xmlNode.OwnerDocument.CreateElement(text3, text4, empty) : xmlNode.OwnerDocument.CreateElement(text4, empty)));
				if (nodeInsertOrder == eNodeInsertOrder.SchemaOrder)
				{
					if (SchemaNodeOrder == null || SchemaNodeOrder.Length == 0)
					{
						nodeInsertOrder = eNodeInsertOrder.Last;
					}
					else
					{
						referenceNode = GetPrependNode(text4, xmlNode);
						nodeInsertOrder = ((referenceNode == null) ? eNodeInsertOrder.Last : eNodeInsertOrder.Before);
					}
				}
				switch (nodeInsertOrder)
				{
				case eNodeInsertOrder.After:
					xmlNode.InsertAfter(xmlNode2, referenceNode);
					referenceNode = null;
					break;
				case eNodeInsertOrder.Before:
					xmlNode.InsertBefore(xmlNode2, referenceNode);
					referenceNode = null;
					break;
				case eNodeInsertOrder.First:
					xmlNode.PrependChild(xmlNode2);
					break;
				case eNodeInsertOrder.Last:
					xmlNode.AppendChild(xmlNode2);
					break;
				}
			}
			xmlNode = xmlNode2;
		}
		return xmlNode;
	}

	private XmlNode GetPrependNode(string nodeName, XmlNode node)
	{
		int nodePos = GetNodePos(nodeName);
		if (nodePos < 0)
		{
			return null;
		}
		XmlNode result = null;
		foreach (XmlNode childNode in node.ChildNodes)
		{
			int nodePos2 = GetNodePos(childNode.Name);
			if (nodePos2 > -1 && nodePos2 > nodePos)
			{
				result = childNode;
				break;
			}
		}
		return result;
	}

	private int GetNodePos(string nodeName)
	{
		int num = nodeName.IndexOf(":");
		if (num > 0)
		{
			nodeName = nodeName.Substring(num + 1, nodeName.Length - (num + 1));
		}
		for (int i = 0; i < _schemaNodeOrder.Length; i++)
		{
			if (nodeName == _schemaNodeOrder[i])
			{
				return i;
			}
		}
		return -1;
	}

	internal void DeleteAllNode(string path)
	{
		string[] array = path.Split('/');
		XmlNode xmlNode = TopNode;
		string[] array2 = array;
		foreach (string xpath in array2)
		{
			xmlNode = xmlNode.SelectSingleNode(xpath, NameSpaceManager);
			if (xmlNode != null)
			{
				if (xmlNode is XmlAttribute)
				{
					(xmlNode as XmlAttribute).OwnerElement.Attributes.Remove(xmlNode as XmlAttribute);
				}
				else
				{
					xmlNode.ParentNode.RemoveChild(xmlNode);
				}
				continue;
			}
			break;
		}
	}

	internal void DeleteNode(string path)
	{
		XmlNode xmlNode = TopNode.SelectSingleNode(path, NameSpaceManager);
		if (xmlNode != null)
		{
			if (xmlNode is XmlAttribute)
			{
				XmlAttribute xmlAttribute = (XmlAttribute)xmlNode;
				xmlAttribute.OwnerElement.Attributes.Remove(xmlAttribute);
			}
			else
			{
				xmlNode.ParentNode.RemoveChild(xmlNode);
			}
		}
	}

	internal void DeleteTopNode()
	{
		TopNode.ParentNode.RemoveChild(TopNode);
	}

	internal void SetXmlNodeString(string path, string value)
	{
		SetXmlNodeString(TopNode, path, value, removeIfBlank: false, insertFirst: false);
	}

	internal void SetXmlNodeString(string path, string value, bool removeIfBlank)
	{
		SetXmlNodeString(TopNode, path, value, removeIfBlank, insertFirst: false);
	}

	internal void SetXmlNodeString(XmlNode node, string path, string value)
	{
		SetXmlNodeString(node, path, value, removeIfBlank: false, insertFirst: false);
	}

	internal void SetXmlNodeString(XmlNode node, string path, string value, bool removeIfBlank)
	{
		SetXmlNodeString(node, path, value, removeIfBlank, insertFirst: false);
	}

	internal void SetXmlNodeString(XmlNode node, string path, string value, bool removeIfBlank, bool insertFirst)
	{
		if (node == null)
		{
			return;
		}
		if (value == "" && removeIfBlank)
		{
			DeleteAllNode(path);
			return;
		}
		XmlNode xmlNode = node.SelectSingleNode(path, NameSpaceManager);
		if (xmlNode == null)
		{
			CreateNode(path, insertFirst);
			xmlNode = node.SelectSingleNode(path, NameSpaceManager);
		}
		xmlNode.InnerText = value;
	}

	internal void SetXmlNodeBool(string path, bool value)
	{
		SetXmlNodeString(TopNode, path, value ? "1" : "0", removeIfBlank: false, insertFirst: false);
	}

	internal void SetXmlNodeBool(string path, bool value, bool removeIf)
	{
		if (value == removeIf)
		{
			XmlNode xmlNode = TopNode.SelectSingleNode(path, NameSpaceManager);
			if (xmlNode != null)
			{
				if (xmlNode is XmlAttribute)
				{
					XmlElement ownerElement = (xmlNode as XmlAttribute).OwnerElement;
					ownerElement.ParentNode.RemoveChild(ownerElement);
				}
				else
				{
					TopNode.RemoveChild(xmlNode);
				}
			}
		}
		else
		{
			SetXmlNodeString(TopNode, path, value ? "1" : "0", removeIfBlank: false, insertFirst: false);
		}
	}

	internal bool ExistNode(string path)
	{
		if (TopNode == null || TopNode.SelectSingleNode(path, NameSpaceManager) == null)
		{
			return false;
		}
		return true;
	}

	internal bool? GetXmlNodeBoolNullable(string path)
	{
		if (string.IsNullOrEmpty(GetXmlNodeString(path)))
		{
			return null;
		}
		return GetXmlNodeBool(path);
	}

	internal bool GetXmlNodeBool(string path)
	{
		return GetXmlNodeBool(path, blankValue: false);
	}

	internal bool GetXmlNodeBool(string path, bool blankValue)
	{
		string xmlNodeString = GetXmlNodeString(path);
		if (xmlNodeString == "1" || xmlNodeString == "-1" || xmlNodeString.Equals("true", StringComparison.OrdinalIgnoreCase))
		{
			return true;
		}
		if (xmlNodeString == "")
		{
			return blankValue;
		}
		return false;
	}

	internal int GetXmlNodeInt(string path)
	{
		if (int.TryParse(GetXmlNodeString(path), NumberStyles.Number, CultureInfo.InvariantCulture, out var result))
		{
			return result;
		}
		return int.MinValue;
	}

	internal int? GetXmlNodeIntNull(string path)
	{
		string xmlNodeString = GetXmlNodeString(path);
		if (xmlNodeString != "" && int.TryParse(xmlNodeString, NumberStyles.Number, CultureInfo.InvariantCulture, out var result))
		{
			return result;
		}
		return null;
	}

	internal decimal GetXmlNodeDecimal(string path)
	{
		if (decimal.TryParse(GetXmlNodeString(path), NumberStyles.Any, CultureInfo.InvariantCulture, out var result))
		{
			return result;
		}
		return 0m;
	}

	internal decimal? GetXmlNodeDecimalNull(string path)
	{
		if (decimal.TryParse(GetXmlNodeString(path), NumberStyles.Any, CultureInfo.InvariantCulture, out var result))
		{
			return result;
		}
		return null;
	}

	internal double? GetXmlNodeDoubleNull(string path)
	{
		string xmlNodeString = GetXmlNodeString(path);
		if (xmlNodeString == "")
		{
			return null;
		}
		if (double.TryParse(xmlNodeString, NumberStyles.Any, CultureInfo.InvariantCulture, out var result))
		{
			return result;
		}
		return null;
	}

	internal double GetXmlNodeDouble(string path)
	{
		string xmlNodeString = GetXmlNodeString(path);
		if (xmlNodeString == "")
		{
			return double.NaN;
		}
		if (double.TryParse(xmlNodeString, NumberStyles.Any, CultureInfo.InvariantCulture, out var result))
		{
			return result;
		}
		return double.NaN;
	}

	internal string GetXmlNodeString(XmlNode node, string path)
	{
		if (node == null)
		{
			return "";
		}
		XmlNode xmlNode = node.SelectSingleNode(path, NameSpaceManager);
		if (xmlNode != null)
		{
			if (xmlNode.NodeType == XmlNodeType.Attribute)
			{
				if (xmlNode.Value == null)
				{
					return "";
				}
				return xmlNode.Value;
			}
			return xmlNode.InnerText;
		}
		return "";
	}

	internal string GetXmlNodeString(string path)
	{
		return GetXmlNodeString(TopNode, path);
	}

	internal static Uri GetNewUri(ZipPackage package, string sUri)
	{
		int id = 1;
		return GetNewUri(package, sUri, ref id);
	}

	internal static Uri GetNewUri(ZipPackage package, string sUri, ref int id)
	{
		Uri uri = new Uri(string.Format(sUri, id), UriKind.Relative);
		while (package.PartExists(uri))
		{
			uri = new Uri(string.Format(sUri, ++id), UriKind.Relative);
		}
		return uri;
	}

	internal void InserAfter(XmlNode parentNode, string beforeNodes, XmlNode newNode)
	{
		string[] array = beforeNodes.Split(',');
		foreach (string xpath in array)
		{
			XmlNode xmlNode = parentNode.SelectSingleNode(xpath, NameSpaceManager);
			if (xmlNode != null)
			{
				parentNode.InsertAfter(newNode, xmlNode);
				return;
			}
		}
		parentNode.InsertAfter(newNode, null);
	}

	internal static void LoadXmlSafe(XmlDocument xmlDoc, Stream stream)
	{
		XmlReaderSettings xmlReaderSettings = new XmlReaderSettings();
		xmlReaderSettings.ProhibitDtd = true;
		XmlReader reader = XmlReader.Create(stream, xmlReaderSettings);
		xmlDoc.Load(reader);
	}

	internal static void LoadXmlSafe(XmlDocument xmlDoc, string xml, Encoding encoding)
	{
		MemoryStream stream = new MemoryStream(encoding.GetBytes(xml));
		LoadXmlSafe(xmlDoc, stream);
	}
}
