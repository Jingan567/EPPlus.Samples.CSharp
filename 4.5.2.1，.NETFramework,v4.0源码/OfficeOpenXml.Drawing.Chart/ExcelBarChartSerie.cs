using System.Xml;

namespace OfficeOpenXml.Drawing.Chart;

public sealed class ExcelBarChartSerie : ExcelChartSerie
{
	private ExcelChartSerieDataLabel _DataLabel;

	private const string INVERTIFNEGATIVE_PATH = "c:invertIfNegative/@val";

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

	internal bool InvertIfNegative
	{
		get
		{
			return GetXmlNodeBool("c:invertIfNegative/@val", blankValue: true);
		}
		set
		{
			SetXmlNodeBool("c:invertIfNegative/@val", value);
		}
	}

	internal ExcelBarChartSerie(ExcelChartSeries chartSeries, XmlNamespaceManager ns, XmlNode node, bool isPivot)
		: base(chartSeries, ns, node, isPivot)
	{
	}
}
