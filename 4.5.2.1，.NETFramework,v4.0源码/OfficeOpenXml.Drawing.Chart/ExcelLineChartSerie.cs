using System;
using System.Drawing;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart;

public sealed class ExcelLineChartSerie : ExcelChartSerie
{
	private ExcelChartSerieDataLabel _DataLabel;

	private const string markerPath = "c:marker/c:symbol/@val";

	private const string smoothPath = "c:smooth/@val";

	private string LINECOLOR_PATH = "c:spPr/a:ln/a:solidFill/a:srgbClr/@val";

	private string MARKERSIZE_PATH = "c:marker/c:size/@val";

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

	public bool Smooth
	{
		get
		{
			return GetXmlNodeBool("c:smooth/@val", blankValue: false);
		}
		set
		{
			SetXmlNodeBool("c:smooth/@val", value);
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
			return Color.FromArgb(Convert.ToInt32(xmlNodeString, 16));
		}
		set
		{
			SetXmlNodeString(LINECOLOR_PATH, value.ToArgb().ToString("X").Substring(2), removeIfBlank: true);
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
			return Color.FromArgb(Convert.ToInt32(xmlNodeString, 16));
		}
		set
		{
			SetXmlNodeString(MARKERLINECOLOR_PATH, value.ToArgb().ToString("X").Substring(2), removeIfBlank: true);
		}
	}

	internal ExcelLineChartSerie(ExcelChartSeries chartSeries, XmlNamespaceManager ns, XmlNode node, bool isPivot)
		: base(chartSeries, ns, node, isPivot)
	{
	}
}
