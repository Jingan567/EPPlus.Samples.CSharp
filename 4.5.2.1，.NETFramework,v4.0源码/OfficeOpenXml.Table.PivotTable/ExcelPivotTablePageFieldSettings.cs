using System.Xml;

namespace OfficeOpenXml.Table.PivotTable;

public class ExcelPivotTablePageFieldSettings : XmlHelper
{
	private ExcelPivotTableField _field;

	internal int Index
	{
		get
		{
			return GetXmlNodeInt("@fld");
		}
		set
		{
			SetXmlNodeString("@fld", value.ToString());
		}
	}

	public string Name
	{
		get
		{
			return GetXmlNodeString("@name");
		}
		set
		{
			SetXmlNodeString("@name", value);
		}
	}

	internal int NumFmtId
	{
		get
		{
			return GetXmlNodeInt("@numFmtId");
		}
		set
		{
			SetXmlNodeString("@numFmtId", value.ToString());
		}
	}

	internal int Hier
	{
		get
		{
			return GetXmlNodeInt("@hier");
		}
		set
		{
			SetXmlNodeString("@hier", value.ToString());
		}
	}

	internal ExcelPivotTablePageFieldSettings(XmlNamespaceManager ns, XmlNode topNode, ExcelPivotTableField field, int index)
		: base(ns, topNode)
	{
		if (GetXmlNodeString("@hier") == "")
		{
			Hier = -1;
		}
		_field = field;
	}
}
