using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart;

public sealed class ExcelPieChartSerie : ExcelChartSerie
{
	private const string explosionPath = "c:explosion/@val";

	private ExcelChartSerieDataLabel _DataLabel;

	public int Explosion
	{
		get
		{
			return GetXmlNodeInt("c:explosion/@val");
		}
		set
		{
			if (value < 0 || value > 400)
			{
				throw new ArgumentOutOfRangeException("Explosion range is 0-400");
			}
			SetXmlNodeString("c:explosion/@val", value.ToString());
		}
	}

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

	internal ExcelPieChartSerie(ExcelChartSeries chartSeries, XmlNamespaceManager ns, XmlNode node, bool isPivot)
		: base(chartSeries, ns, node, isPivot)
	{
	}
}
