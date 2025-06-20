using System.Xml;

namespace OfficeOpenXml.Style.XmlAccess;

public abstract class StyleXmlHelper : XmlHelper
{
	internal long useCnt;

	internal int newID = int.MinValue;

	internal abstract string Id { get; }

	internal StyleXmlHelper(XmlNamespaceManager nameSpaceManager)
		: base(nameSpaceManager)
	{
	}

	internal StyleXmlHelper(XmlNamespaceManager nameSpaceManager, XmlNode topNode)
		: base(nameSpaceManager, topNode)
	{
	}

	internal abstract XmlNode CreateXmlNode(XmlNode top);

	protected bool GetBoolValue(XmlNode topNode, string path)
	{
		XmlNode xmlNode = topNode.SelectSingleNode(path, base.NameSpaceManager);
		if (xmlNode is XmlAttribute)
		{
			return xmlNode.Value != "0";
		}
		if (xmlNode != null && ((xmlNode.Attributes["val"] != null && xmlNode.Attributes["val"].Value != "0") || xmlNode.Attributes["val"] == null))
		{
			return true;
		}
		return false;
	}
}
