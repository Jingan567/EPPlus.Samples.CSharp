using System;
using System.Xml;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Table.PivotTable;

namespace OfficeOpenXml.Drawing.Chart;

public class ExcelPieChart : ExcelChart
{
	private ExcelChartDataLabel _DataLabel;

	public ExcelChartDataLabel DataLabel
	{
		get
		{
			if (_DataLabel == null)
			{
				_DataLabel = new ExcelChartDataLabel(base.NameSpaceManager, base.ChartNode);
			}
			return _DataLabel;
		}
	}

	internal ExcelPieChart(ExcelDrawings drawings, XmlNode node, eChartType type, bool isPivot)
		: base(drawings, node, type, isPivot)
	{
	}

	internal ExcelPieChart(ExcelDrawings drawings, XmlNode node, eChartType type, ExcelChart topChart, ExcelPivotTable PivotTableSource)
		: base(drawings, node, type, topChart, PivotTableSource)
	{
	}

	internal ExcelPieChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode)
		: base(drawings, node, uriChart, part, chartXml, chartNode)
	{
	}

	internal ExcelPieChart(ExcelChart topChart, XmlNode chartNode)
		: base(topChart, chartNode)
	{
	}

	internal override eChartType GetChartType(string name)
	{
		if (name == "pieChart")
		{
			if (Series.Count > 0 && ((ExcelPieChartSerie)Series[0]).Explosion > 0)
			{
				return eChartType.PieExploded;
			}
			return eChartType.Pie;
		}
		if (name == "pie3DChart")
		{
			if (Series.Count > 0 && ((ExcelPieChartSerie)Series[0]).Explosion > 0)
			{
				return eChartType.PieExploded3D;
			}
			return eChartType.Pie3D;
		}
		return base.GetChartType(name);
	}
}
