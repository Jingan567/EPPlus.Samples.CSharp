using System;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable;

public class ExcelPivotTableFieldDateGroup : ExcelPivotTableFieldGroup
{
	private const string groupByPath = "d:fieldGroup/d:rangePr/@groupBy";

	public eDateGroupBy GroupBy
	{
		get
		{
			string xmlNodeString = GetXmlNodeString("d:fieldGroup/d:rangePr/@groupBy");
			if (xmlNodeString != "")
			{
				return (eDateGroupBy)Enum.Parse(typeof(eDateGroupBy), xmlNodeString, ignoreCase: true);
			}
			throw new Exception("Invalid date Groupby");
		}
		private set
		{
			SetXmlNodeString("d:fieldGroup/d:rangePr/@groupBy", value.ToString().ToLower(CultureInfo.InvariantCulture));
		}
	}

	public bool AutoStart => GetXmlNodeBool("@autoStart", blankValue: false);

	public bool AutoEnd => GetXmlNodeBool("@autoStart", blankValue: false);

	internal ExcelPivotTableFieldDateGroup(XmlNamespaceManager ns, XmlNode topNode)
		: base(ns, topNode)
	{
	}
}
