using System;
using System.Globalization;
using System.Xml;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Table.PivotTable;

namespace OfficeOpenXml.Drawing.Chart;

public class ExcelRadarChart : ExcelChart
{
	private string STYLE_PATH = "c:radarStyle/@val";

	private ExcelChartDataLabel _DataLabel;

	public eRadarStyle RadarStyle
	{
		get
		{
			string xmlNodeString = _chartXmlHelper.GetXmlNodeString(STYLE_PATH);
			if (string.IsNullOrEmpty(xmlNodeString))
			{
				return eRadarStyle.Standard;
			}
			return (eRadarStyle)Enum.Parse(typeof(eRadarStyle), xmlNodeString, ignoreCase: true);
		}
		set
		{
			_chartXmlHelper.SetXmlNodeString(STYLE_PATH, value.ToString().ToLower(CultureInfo.InvariantCulture));
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

	internal ExcelRadarChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode)
		: base(drawings, node, uriChart, part, chartXml, chartNode)
	{
		SetTypeProperties();
	}

	internal ExcelRadarChart(ExcelChart topChart, XmlNode chartNode)
		: base(topChart, chartNode)
	{
		SetTypeProperties();
	}

	internal ExcelRadarChart(ExcelDrawings drawings, XmlNode node, eChartType type, ExcelChart topChart, ExcelPivotTable PivotTableSource)
		: base(drawings, node, type, topChart, PivotTableSource)
	{
		SetTypeProperties();
	}

	private void SetTypeProperties()
	{
		if (base.ChartType == eChartType.RadarFilled)
		{
			RadarStyle = eRadarStyle.Filled;
		}
		else if (base.ChartType == eChartType.RadarMarkers)
		{
			RadarStyle = eRadarStyle.Marker;
		}
		else
		{
			RadarStyle = eRadarStyle.Standard;
		}
	}

	internal override eChartType GetChartType(string name)
	{
		if (RadarStyle == eRadarStyle.Filled)
		{
			return eChartType.RadarFilled;
		}
		if (RadarStyle == eRadarStyle.Marker)
		{
			return eChartType.RadarMarkers;
		}
		return eChartType.Radar;
	}
}
