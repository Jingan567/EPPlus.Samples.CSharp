using System.Xml;

namespace OfficeOpenXml.Drawing.Chart;

public sealed class ExcelBubbleChartSeries : ExcelChartSeries
{
	internal ExcelBubbleChartSeries(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, bool isPivot)
		: base(chart, ns, node, isPivot)
	{
	}

	public ExcelChartSerie Add(ExcelRangeBase Serie, ExcelRangeBase XSerie, ExcelRangeBase BubbleSize)
	{
		return AddSeries(Serie.FullAddressAbsolute, XSerie.FullAddressAbsolute, BubbleSize.FullAddressAbsolute);
	}

	public ExcelChartSerie Add(string SerieAddress, string XSerieAddress, string BubbleSizeAddress)
	{
		return AddSeries(SerieAddress, XSerieAddress, BubbleSizeAddress);
	}
}
