using System;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart;

public sealed class ExcelRadarChartSerie : ExcelChartSerie
{
	private ExcelChartSerieDataLabel _DataLabel;

	private const string markerPath = "c:marker/c:symbol/@val";

	private const string MARKERSIZE_PATH = "c:marker/c:size/@val";

	public ExcelChartSerieDataLabel DataLabel
	{
		get
		{
			if (_DataLabel == null)
			{
				_DataLabel = new ExcelChartSerieDataLabel(_ns, _node);
			}
			return _DataLabel;
		}
	}

	public eMarkerStyle Marker
	{
		get
		{
			string xmlNodeString = GetXmlNodeString("c:marker/c:symbol/@val");
			if (xmlNodeString == "" || xmlNodeString == "none")
			{
				return eMarkerStyle.None;
			}
			return (eMarkerStyle)Enum.Parse(typeof(eMarkerStyle), xmlNodeString, ignoreCase: true);
		}
		internal set
		{
			SetXmlNodeString("c:marker/c:symbol/@val", value.ToString().ToLower(CultureInfo.InvariantCulture));
		}
	}

	public int MarkerSize
	{
		get
		{
			return GetXmlNodeInt("c:marker/c:size/@val");
		}
		set
		{
			if (value < 2 && value > 72)
			{
				throw new ArgumentOutOfRangeException("MarkerSize out of range. Range from 2-72 allowed.");
			}
			SetXmlNodeString("c:marker/c:size/@val", value.ToString(CultureInfo.InvariantCulture));
		}
	}

	internal ExcelRadarChartSerie(ExcelChartSeries chartSeries, XmlNamespaceManager ns, XmlNode node, bool isPivot)
		: base(chartSeries, ns, node, isPivot)
	{
		if (chartSeries.Chart.ChartType == eChartType.Radar)
		{
			Marker = eMarkerStyle.None;
		}
	}
}
