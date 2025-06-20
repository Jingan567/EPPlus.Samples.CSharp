using System.Xml;

namespace OfficeOpenXml.Table.PivotTable;

public class ExcelPivotTableFieldGroup : XmlHelper
{
	internal ExcelPivotTableFieldGroup(XmlNamespaceManager ns, XmlNode topNode)
		: base(ns, topNode)
	{
	}
}
