using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart;

public sealed class ExcelScatterChartSerie : ExcelChartSerie
{
	private ExcelChartSerieDataLabel _DataLabel;

	private const string smoothPath = "c:smooth/@val";

	private const string markerPath = "c:marker/c:symbol/@val";

	private string LINECOLOR_PATH = "c:spPr/a:ln/a:solidFill/a:srgbClr/@val";

	private string MARKERSIZE_PATH = "c:marker/c:size/@val";

	private string MARKERCOLOR_PATH = "c:marker/c:spPr/a:solidFill/a:srgbClr/@val";

	private string LINEWIDTH_PATH = "c:spPr/a:ln/@w";

	private string MARKERLINECOLOR_PATH = "c:marker/c:spPr/a:ln/a:solidFill/a:srgbClr/@val";

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

	public int Smooth
	{
		get
		{
			return GetXmlNodeInt("c:smooth/@val");
		}
		internal set
		{
			SetXmlNodeString("c:smooth/@val", value.ToString());
		}
	}

	public eMarkerStyle Marker
	{
		get
		{
			string xmlNodeString = GetXmlNodeString("c:marker/c:symbol/@val");
			if (xmlNodeString == "")
			{
				return eMarkerStyle.None;
			}
			return (eMarkerStyle)Enum.Parse(typeof(eMarkerStyle), xmlNodeString, ignoreCase: true);
		}
		set
		{
			SetXmlNodeString("c:marker/c:symbol/@val", value.ToString().ToLower(CultureInfo.InvariantCulture));
		}
	}

	public Color LineColor
	{
		get
		{
			string xmlNodeString = GetXmlNodeString(LINECOLOR_PATH);
			if (xmlNodeString == "")
			{
				return Color.Black;
			}
			Color color = Color.FromArgb(Convert.ToInt32(xmlNodeString, 16));
			int alphaChannel = getAlphaChannel(LINECOLOR_PATH);
			if (alphaChannel != 255)
			{
				color = Color.FromArgb(alphaChannel, color);
			}
			return color;
		}
		set
		{
			SetXmlNodeString(LINECOLOR_PATH, value.ToArgb().ToString("X8").Substring(2), removeIfBlank: true);
			setAlphaChannel(value, LINECOLOR_PATH);
		}
	}

	public int MarkerSize
	{
		get
		{
			if (GetXmlNodeString(MARKERSIZE_PATH) == "")
			{
				return 5;
			}
			return int.Parse(GetXmlNodeString(MARKERSIZE_PATH));
		}
		set
		{
			int val = value;
			val = Math.Max(2, val);
			val = Math.Min(72, val);
			SetXmlNodeString(MARKERSIZE_PATH, val.ToString(), removeIfBlank: true);
		}
	}

	public Color MarkerColor
	{
		get
		{
			string xmlNodeString = GetXmlNodeString(MARKERCOLOR_PATH);
			if (xmlNodeString == "")
			{
				return Color.Black;
			}
			Color color = Color.FromArgb(Convert.ToInt32(xmlNodeString, 16));
			int alphaChannel = getAlphaChannel(MARKERCOLOR_PATH);
			if (alphaChannel != 255)
			{
				color = Color.FromArgb(alphaChannel, color);
			}
			return color;
		}
		set
		{
			SetXmlNodeString(MARKERCOLOR_PATH, value.ToArgb().ToString("X8").Substring(2), removeIfBlank: true);
			setAlphaChannel(value, MARKERCOLOR_PATH);
		}
	}

	public double LineWidth
	{
		get
		{
			if (GetXmlNodeString(LINEWIDTH_PATH) == "")
			{
				return 2.25;
			}
			return double.Parse(GetXmlNodeString(LINEWIDTH_PATH)) / 12700.0;
		}
		set
		{
			SetXmlNodeString(LINEWIDTH_PATH, ((int)(12700.0 * value)).ToString(), removeIfBlank: true);
		}
	}

	public Color MarkerLineColor
	{
		get
		{
			string xmlNodeString = GetXmlNodeString(MARKERLINECOLOR_PATH);
			if (xmlNodeString == "")
			{
				return Color.Black;
			}
			Color color = Color.FromArgb(Convert.ToInt32(xmlNodeString, 16));
			int alphaChannel = getAlphaChannel(MARKERLINECOLOR_PATH);
			if (alphaChannel != 255)
			{
				color = Color.FromArgb(alphaChannel, color);
			}
			return color;
		}
		set
		{
			SetXmlNodeString(MARKERLINECOLOR_PATH, value.ToArgb().ToString("X8").Substring(2), removeIfBlank: true);
			setAlphaChannel(value, MARKERLINECOLOR_PATH);
		}
	}

	internal ExcelScatterChartSerie(ExcelChartSeries chartSeries, XmlNamespaceManager ns, XmlNode node, bool isPivot)
		: base(chartSeries, ns, node, isPivot)
	{
		if (chartSeries.Chart.ChartType == eChartType.XYScatterLines || chartSeries.Chart.ChartType == eChartType.XYScatterSmooth)
		{
			Marker = eMarkerStyle.Square;
		}
		if (chartSeries.Chart.ChartType == eChartType.XYScatterSmooth || chartSeries.Chart.ChartType == eChartType.XYScatterSmoothNoMarkers)
		{
			Smooth = 1;
		}
		else if (chartSeries.Chart.ChartType == eChartType.XYScatterLines || chartSeries.Chart.ChartType == eChartType.XYScatterLinesNoMarkers || chartSeries.Chart.ChartType == eChartType.XYScatter)
		{
			Smooth = 0;
		}
	}

	private void setAlphaChannel(Color c, string xPath)
	{
		if (c.A != byte.MaxValue)
		{
			string text = xPath4Alpha(xPath);
			if (text.Length > 0)
			{
				string value = ((c.A != 0) ? ((100 - c.A) * 1000) : 0).ToString();
				SetXmlNodeString(text, value, removeIfBlank: true);
			}
		}
	}

	private int getAlphaChannel(string xPath)
	{
		int result = 255;
		string text = xPath4Alpha(xPath);
		if (text.Length > 0)
		{
			int result2 = 0;
			if (int.TryParse(GetXmlNodeString(text), NumberStyles.Any, CultureInfo.InvariantCulture, out result2))
			{
				result = ((result2 != 0) ? (100 - result2 / 1000) : 0);
			}
		}
		return result;
	}

	private string xPath4Alpha(string xPath)
	{
		string empty = string.Empty;
		if (xPath.EndsWith("@val"))
		{
			xPath = xPath.Substring(0, xPath.IndexOf("@val"));
		}
		if (xPath.EndsWith("/"))
		{
			xPath = xPath.Substring(0, xPath.Length - 1);
		}
		if (new List<string> { "a:prstClr", "a:hslClr", "a:schemeClr", "a:sysClr", "a:scrgbClr", "a:srgbClr" }.Find((string cd) => xPath.EndsWith(cd, StringComparison.Ordinal)) != null)
		{
			return xPath + "/a:alpha/@val";
		}
		throw new InvalidOperationException("alpha-values can only set to Colors");
	}
}
