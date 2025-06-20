using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable;

public class ExcelPivotTableFieldNumericGroup : ExcelPivotTableFieldGroup
{
	private const string startPath = "d:fieldGroup/d:rangePr/@startNum";

	private const string endPath = "d:fieldGroup/d:rangePr/@endNum";

	private const string groupIntervalPath = "d:fieldGroup/d:rangePr/@groupInterval";

	public double Start
	{
		get
		{
			return GetXmlNodeDoubleNull("d:fieldGroup/d:rangePr/@startNum").Value;
		}
		private set
		{
			SetXmlNodeString("d:fieldGroup/d:rangePr/@startNum", value.ToString(CultureInfo.InvariantCulture));
		}
	}

	public double End
	{
		get
		{
			return GetXmlNodeDoubleNull("d:fieldGroup/d:rangePr/@endNum").Value;
		}
		private set
		{
			SetXmlNodeString("d:fieldGroup/d:rangePr/@endNum", value.ToString(CultureInfo.InvariantCulture));
		}
	}

	public double Interval
	{
		get
		{
			return GetXmlNodeDoubleNull("d:fieldGroup/d:rangePr/@groupInterval").Value;
		}
		private set
		{
			SetXmlNodeString("d:fieldGroup/d:rangePr/@groupInterval", value.ToString(CultureInfo.InvariantCulture));
		}
	}

	internal ExcelPivotTableFieldNumericGroup(XmlNamespaceManager ns, XmlNode topNode)
		: base(ns, topNode)
	{
	}
}
