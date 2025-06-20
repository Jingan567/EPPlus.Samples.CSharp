using System.Xml;

namespace OfficeOpenXml;

internal static class XmlHelperFactory
{
	internal static XmlHelper Create(XmlNamespaceManager namespaceManager)
	{
		return new XmlHelperInstance(namespaceManager);
	}

	internal static XmlHelper Create(XmlNamespaceManager namespaceManager, XmlNode topNode)
	{
		return new XmlHelperInstance(namespaceManager, topNode);
	}
}
