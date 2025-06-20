using System;
using System.Xml;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Table.PivotTable;

namespace OfficeOpenXml.Drawing.Chart;

public class ExcelLineChart : ExcelChart
{
	private string MARKER_PATH = "c:marker/@val";

	private string SMOOTH_PATH = "c:smooth/@val";

	private ExcelChartDataLabel _DataLabel;

	public bool Marker
	{
		get
		{
			return _chartXmlHelper.GetXmlNodeBool(MARKER_PATH, blankValue: false);
		}
		set
		{
			_chartXmlHelper.SetXmlNodeBool(MARKER_PATH, value, removeIf: false);
		}
	}

	public bool Smooth
	{
		get
		{
			return _chartXmlHelper.GetXmlNodeBool(SMOOTH_PATH, blankValue: false);
		}
		set
		{
			_chartXmlHelper.SetXmlNodeBool(SMOOTH_PATH, value);
		}
	}

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

	internal ExcelLineChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode)
		: base(drawings, node, uriChart, part, chartXml, chartNode)
	{
	}

	internal ExcelLineChart(ExcelChart topChart, XmlNode chartNode)
		: base(topChart, chartNode)
	{
	}

	internal ExcelLineChart(ExcelDrawings drawings, XmlNode node, eChartType type, ExcelChart topChart, ExcelPivotTable PivotTableSource)
		: base(drawings, node, type, topChart, PivotTableSource)
	{
		Smooth = false;
	}

	internal override eChartType GetChartType(string name)
	{
		if (name == "lineChart")
		{
			if (Marker)
			{
				if (base.Grouping == eGrouping.Stacked)
				{
					return eChartType.LineMarkersStacked;
				}
				if (base.Grouping == eGrouping.PercentStacked)
				{
					return eChartType.LineMarkersStacked100;
				}
				return eChartType.LineMarkers;
			}
			if (base.Grouping == eGrouping.Stacked)
			{
				return eChartType.LineStacked;
			}
			if (base.Grouping == eGrouping.PercentStacked)
			{
				return eChartType.LineStacked100;
			}
			return eChartType.Line;
		}
		if (name == "line3DChart")
		{
			return eChartType.Line3D;
		}
		return base.GetChartType(name);
	}
}
