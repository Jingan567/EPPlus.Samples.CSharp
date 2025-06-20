using System;
using System.Globalization;
using System.Xml;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Table.PivotTable;

namespace OfficeOpenXml.Drawing.Chart;

public class ExcelDoughnutChart : ExcelPieChart
{
	private string _firstSliceAngPath = "c:firstSliceAng/@val";

	private string _holeSizePath = "c:holeSize/@val";

	public decimal FirstSliceAngle
	{
		get
		{
			return _chartXmlHelper.GetXmlNodeDecimal(_firstSliceAngPath);
		}
		internal set
		{
			_chartXmlHelper.SetXmlNodeString(_firstSliceAngPath, value.ToString(CultureInfo.InvariantCulture));
		}
	}

	public decimal HoleSize
	{
		get
		{
			return _chartXmlHelper.GetXmlNodeDecimal(_holeSizePath);
		}
		internal set
		{
			_chartXmlHelper.SetXmlNodeString(_holeSizePath, value.ToString(CultureInfo.InvariantCulture));
		}
	}

	internal ExcelDoughnutChart(ExcelDrawings drawings, XmlNode node, eChartType type, bool isPivot)
		: base(drawings, node, type, isPivot)
	{
	}

	internal ExcelDoughnutChart(ExcelDrawings drawings, XmlNode node, eChartType type, ExcelChart topChart, ExcelPivotTable PivotTableSource)
		: base(drawings, node, type, topChart, PivotTableSource)
	{
	}

	internal ExcelDoughnutChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode)
		: base(drawings, node, uriChart, part, chartXml, chartNode)
	{
	}

	internal ExcelDoughnutChart(ExcelChart topChart, XmlNode chartNode)
		: base(topChart, chartNode)
	{
	}

	internal override eChartType GetChartType(string name)
	{
		if (name == "doughnutChart")
		{
			if (((ExcelPieChartSerie)Series[0]).Explosion > 0)
			{
				return eChartType.DoughnutExploded;
			}
			return eChartType.Doughnut;
		}
		return base.GetChartType(name);
	}
}
