using System.Xml;

namespace OfficeOpenXml.Drawing.Chart;

public sealed class ExcelSurfaceChartSerie : ExcelChartSerie
{
	internal ExcelSurfaceChartSerie(ExcelChartSeries chartSeries, XmlNamespaceManager ns, XmlNode node, bool isPivot)
		: base(chartSeries, ns, node, isPivot)
	{
	}
}
