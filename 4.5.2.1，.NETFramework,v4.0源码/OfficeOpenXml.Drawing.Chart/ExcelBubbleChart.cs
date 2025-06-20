using System;
using System.Globalization;
using System.Xml;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Table.PivotTable;

namespace OfficeOpenXml.Drawing.Chart;

public sealed class ExcelBubbleChart : ExcelChart
{
	private string BUBBLESCALE_PATH = "c:bubbleScale/@val";

	private string SHOWNEGBUBBLES_PATH = "c:showNegBubbles/@val";

	private string BUBBLE3D_PATH = "c:bubble3D/@val";

	private string SIZEREPRESENTS_PATH = "c:sizeRepresents/@val";

	public int BubbleScale
	{
		get
		{
			return _chartXmlHelper.GetXmlNodeInt(BUBBLESCALE_PATH);
		}
		set
		{
			if (value < 0 && value > 300)
			{
				throw new ArgumentOutOfRangeException("Bubblescale out of range. 0-300 allowed");
			}
			_chartXmlHelper.SetXmlNodeString(BUBBLESCALE_PATH, value.ToString());
		}
	}

	public bool ShowNegativeBubbles
	{
		get
		{
			return _chartXmlHelper.GetXmlNodeBool(SHOWNEGBUBBLES_PATH);
		}
		set
		{
			_chartXmlHelper.SetXmlNodeBool(BUBBLESCALE_PATH, value, removeIf: true);
		}
	}

	public bool Bubble3D
	{
		get
		{
			return _chartXmlHelper.GetXmlNodeBool(BUBBLE3D_PATH);
		}
		set
		{
			_chartXmlHelper.SetXmlNodeBool(BUBBLE3D_PATH, value);
			base.ChartType = (value ? eChartType.Bubble3DEffect : eChartType.Bubble);
		}
	}

	public eSizeRepresents SizeRepresents
	{
		get
		{
			if (_chartXmlHelper.GetXmlNodeString(SIZEREPRESENTS_PATH).ToLower(CultureInfo.InvariantCulture) == "w")
			{
				return eSizeRepresents.Width;
			}
			return eSizeRepresents.Area;
		}
		set
		{
			_chartXmlHelper.SetXmlNodeString(SIZEREPRESENTS_PATH, (value == eSizeRepresents.Width) ? "w" : "area");
		}
	}

	public new ExcelBubbleChartSeries Series => (ExcelBubbleChartSeries)_chartSeries;

	internal ExcelBubbleChart(ExcelDrawings drawings, XmlNode node, eChartType type, ExcelChart topChart, ExcelPivotTable PivotTableSource)
		: base(drawings, node, type, topChart, PivotTableSource)
	{
		ShowNegativeBubbles = false;
		BubbleScale = 100;
		_chartSeries = new ExcelBubbleChartSeries(this, drawings.NameSpaceManager, _chartNode, PivotTableSource != null);
	}

	internal ExcelBubbleChart(ExcelDrawings drawings, XmlNode node, eChartType type, bool isPivot)
		: base(drawings, node, type, isPivot)
	{
		_chartSeries = new ExcelBubbleChartSeries(this, drawings.NameSpaceManager, _chartNode, isPivot);
	}

	internal ExcelBubbleChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode)
		: base(drawings, node, uriChart, part, chartXml, chartNode)
	{
		_chartSeries = new ExcelBubbleChartSeries(this, _drawings.NameSpaceManager, _chartNode, isPivot: false);
	}

	internal ExcelBubbleChart(ExcelChart topChart, XmlNode chartNode)
		: base(topChart, chartNode)
	{
		_chartSeries = new ExcelBubbleChartSeries(this, _drawings.NameSpaceManager, _chartNode, isPivot: false);
	}

	internal override eChartType GetChartType(string name)
	{
		if (Bubble3D)
		{
			return eChartType.Bubble3DEffect;
		}
		return eChartType.Bubble;
	}
}
