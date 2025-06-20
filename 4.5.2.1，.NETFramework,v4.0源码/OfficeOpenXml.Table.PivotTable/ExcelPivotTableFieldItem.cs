using System;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable;

public class ExcelPivotTableFieldItem : XmlHelper
{
	private ExcelPivotTableField _field;

	public string Text
	{
		get
		{
			return GetXmlNodeString("@n");
		}
		set
		{
			if (string.IsNullOrEmpty(value))
			{
				DeleteNode("@n");
				return;
			}
			foreach (ExcelPivotTableFieldItem item in _field.Items)
			{
				if (item.Text == value)
				{
					throw new ArgumentException("Duplicate Text");
				}
			}
			SetXmlNodeString("@n", value);
		}
	}

	internal int X => GetXmlNodeInt("@x");

	internal string T => GetXmlNodeString("@t");

	internal ExcelPivotTableFieldItem(XmlNamespaceManager ns, XmlNode topNode, ExcelPivotTableField field)
		: base(ns, topNode)
	{
		_field = field;
	}
}
