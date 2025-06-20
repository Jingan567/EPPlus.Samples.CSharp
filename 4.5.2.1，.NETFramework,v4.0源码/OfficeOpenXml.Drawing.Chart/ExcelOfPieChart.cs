using System;
using System.Globalization;
using System.Xml;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Table.PivotTable;

namespace OfficeOpenXml.Drawing.Chart;

public class ExcelOfPieChart : ExcelPieChart
{
	private const string pieTypePath = "c:ofPieType/@val";

	private string _gapWidthPath = "c:gapWidth/@val";

	public ePieType OfPieType
	{
		get
		{
			if (_chartXmlHelper.GetXmlNodeString("c:ofPieType/@val") == "bar")
			{
				return ePieType.Bar;
			}
			return ePieType.Pie;
		}
		internal set
		{
			_chartXmlHelper.CreateNode("c:ofPieType/@val", insertFirst: true);
			_chartXmlHelper.SetXmlNodeString("c:ofPieType/@val", (value == ePieType.Bar) ? "bar" : "pie");
		}
	}

	public int GapWidth
	{
		get
		{
			return _chartXmlHelper.GetXmlNodeInt(_gapWidthPath);
		}
		set
		{
			_chartXmlHelper.SetXmlNodeString(_gapWidthPath, value.ToString(CultureInfo.InvariantCulture));
		}
	}

	internal ExcelOfPieChart(ExcelDrawings drawings, XmlNode node, eChartType type, bool isPivot)
		: base(drawings, node, type, isPivot)
	{
		SetTypeProperties();
	}

	internal ExcelOfPieChart(ExcelDrawings drawings, XmlNode node, eChartType type, ExcelChart topChart, ExcelPivotTable PivotTableSource)
		: base(drawings, node, type, topChart, PivotTableSource)
	{
		SetTypeProperties();
	}

	internal ExcelOfPieChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode)
		: base(drawings, node, uriChart, part, chartXml, chartNode)
	{
		SetTypeProperties();
	}

	private void SetTypeProperties()
	{
		if (base.ChartType == eChartType.BarOfPie)
		{
			OfPieType = ePieType.Bar;
		}
		else
		{
			OfPieType = ePieType.Pie;
		}
	}

	internal override eChartType GetChartType(string name)
	{
		if (name == "ofPieChart")
		{
			if (OfPieType == ePieType.Bar)
			{
				return eChartType.BarOfPie;
			}
			return eChartType.PieOfPie;
		}
		return base.GetChartType(name);
	}
}
