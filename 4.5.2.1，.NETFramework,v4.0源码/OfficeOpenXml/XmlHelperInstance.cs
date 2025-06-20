using System.Xml;

namespace OfficeOpenXml;

internal class XmlHelperInstance : XmlHelper
{
	internal XmlHelperInstance(XmlNamespaceManager namespaceManager)
		: base(namespaceManager)
	{
	}

	internal XmlHelperInstance(XmlNamespaceManager namespaceManager, XmlNode topNode)
		: base(namespaceManager, topNode)
	{
	}
}
