using System;
using System.Xml;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Table.PivotTable;

namespace OfficeOpenXml.Drawing.Chart;

public sealed class ExcelScatterChart : ExcelChart
{
	private string _scatterTypePath = "c:scatterStyle/@val";

	private string MARKER_PATH = "c:marker/@val";

	public eScatterStyle ScatterStyle
	{
		get
		{
			return GetScatterEnum(_chartXmlHelper.GetXmlNodeString(_scatterTypePath));
		}
		internal set
		{
			_chartXmlHelper.CreateNode(_scatterTypePath, insertFirst: true);
			_chartXmlHelper.SetXmlNodeString(_scatterTypePath, GetScatterText(value));
		}
	}

	public bool Marker
	{
		get
		{
			return GetXmlNodeBool(MARKER_PATH, blankValue: false);
		}
		set
		{
			SetXmlNodeBool(MARKER_PATH, value, removeIf: false);
		}
	}

	internal ExcelScatterChart(ExcelDrawings drawings, XmlNode node, eChartType type, ExcelChart topChart, ExcelPivotTable PivotTableSource)
		: base(drawings, node, type, topChart, PivotTableSource)
	{
		SetTypeProperties();
	}

	internal ExcelScatterChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode)
		: base(drawings, node, uriChart, part, chartXml, chartNode)
	{
		SetTypeProperties();
	}

	internal ExcelScatterChart(ExcelChart topChart, XmlNode chartNode)
		: base(topChart, chartNode)
	{
		SetTypeProperties();
	}

	private void SetTypeProperties()
	{
		if (base.ChartType == eChartType.XYScatter || base.ChartType == eChartType.XYScatterLines || base.ChartType == eChartType.XYScatterLinesNoMarkers)
		{
			ScatterStyle = eScatterStyle.LineMarker;
		}
		else if (base.ChartType == eChartType.XYScatterSmooth || base.ChartType == eChartType.XYScatterSmoothNoMarkers)
		{
			ScatterStyle = eScatterStyle.SmoothMarker;
		}
	}

	private eScatterStyle GetScatterEnum(string text)
	{
		if (text == "smoothMarker")
		{
			return eScatterStyle.SmoothMarker;
		}
		return eScatterStyle.LineMarker;
	}

	private string GetScatterText(eScatterStyle shatterStyle)
	{
		if (shatterStyle == eScatterStyle.SmoothMarker)
		{
			return "smoothMarker";
		}
		return "lineMarker";
	}

	internal override eChartType GetChartType(string name)
	{
		if (name == "scatterChart")
		{
			if (ScatterStyle == eScatterStyle.LineMarker)
			{
				if (((ExcelScatterChartSerie)Series[0]).Marker == eMarkerStyle.None)
				{
					return eChartType.XYScatterLinesNoMarkers;
				}
				if (ExistNode("c:ser/c:spPr/a:ln/noFill"))
				{
					return eChartType.XYScatter;
				}
				return eChartType.XYScatterLines;
			}
			if (ScatterStyle == eScatterStyle.SmoothMarker)
			{
				if (((ExcelScatterChartSerie)Series[0]).Marker == eMarkerStyle.None)
				{
					return eChartType.XYScatterSmoothNoMarkers;
				}
				return eChartType.XYScatterSmooth;
			}
		}
		return base.GetChartType(name);
	}
}
