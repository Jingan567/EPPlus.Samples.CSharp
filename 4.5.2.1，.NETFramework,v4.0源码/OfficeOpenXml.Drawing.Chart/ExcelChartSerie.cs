using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart;

public class ExcelChartSerie : XmlHelper
{
	internal ExcelChartSeries _chartSeries;

	protected XmlNode _node;

	protected XmlNamespaceManager _ns;

	private const string headerPath = "c:tx/c:v";

	private const string headerAddressPath = "c:tx/c:strRef/c:f";

	private string _seriesTopPath;

	private string _seriesPath = "{0}/c:numRef/c:f";

	private string _xSeriesTopPath;

	private string _xSeriesPath = "{0}/{1}/c:f";

	private ExcelChartTrendlineCollection _trendLines;

	private ExcelDrawingFill _fill;

	private ExcelDrawingBorder _border;

	public string Header
	{
		get
		{
			return GetXmlNodeString("c:tx/c:v");
		}
		set
		{
			Cleartx();
			SetXmlNodeString("c:tx/c:v", value);
		}
	}

	public ExcelAddressBase HeaderAddress
	{
		get
		{
			string xmlNodeString = GetXmlNodeString("c:tx/c:strRef/c:f");
			if (xmlNodeString == "")
			{
				return null;
			}
			return new ExcelAddressBase(xmlNodeString);
		}
		set
		{
			if ((value._fromCol != value._toCol && value._fromRow != value._toRow) || value.Addresses != null)
			{
				throw new ArgumentException("Address must be a row, column or single cell");
			}
			Cleartx();
			SetXmlNodeString("c:tx/c:strRef/c:f", ExcelCellBase.GetFullAddress(value.WorkSheet, value.Address));
			SetXmlNodeString("c:tx/c:strRef/c:strCache/c:ptCount/@val", "0");
		}
	}

	public virtual string Series
	{
		get
		{
			return GetXmlNodeString(_seriesPath);
		}
		set
		{
			CreateNode(_seriesPath, insertFirst: true);
			SetXmlNodeString(_seriesPath, ExcelCellBase.GetFullAddress(_chartSeries.Chart.WorkSheet.Name, value));
			if (_chartSeries.Chart.PivotTableSource != null)
			{
				XmlNode xmlNode = base.TopNode.SelectSingleNode($"{_seriesTopPath}/c:numRef/c:numCache", _ns);
				xmlNode?.ParentNode.RemoveChild(xmlNode);
				SetXmlNodeString($"{_seriesTopPath}/c:numRef/c:numCache", "General");
			}
			XmlNode xmlNode2 = base.TopNode.SelectSingleNode($"{_seriesTopPath}/c:numLit", _ns);
			xmlNode2?.ParentNode.RemoveChild(xmlNode2);
		}
	}

	public virtual string XSeries
	{
		get
		{
			return GetXmlNodeString(_xSeriesPath);
		}
		set
		{
			CreateNode(_xSeriesPath, insertFirst: true);
			SetXmlNodeString(_xSeriesPath, ExcelCellBase.GetFullAddress(_chartSeries.Chart.WorkSheet.Name, value));
			if (_xSeriesPath.IndexOf("c:numRef") > 0)
			{
				XmlNode xmlNode = base.TopNode.SelectSingleNode($"{_xSeriesTopPath}/c:numRef/c:numCache", _ns);
				xmlNode?.ParentNode.RemoveChild(xmlNode);
				XmlNode xmlNode2 = base.TopNode.SelectSingleNode($"{_xSeriesTopPath}/c:numLit", _ns);
				xmlNode2?.ParentNode.RemoveChild(xmlNode2);
			}
			else
			{
				XmlNode xmlNode3 = base.TopNode.SelectSingleNode($"{_xSeriesTopPath}/c:strRef/c:strCache", _ns);
				xmlNode3?.ParentNode.RemoveChild(xmlNode3);
				XmlNode xmlNode4 = base.TopNode.SelectSingleNode($"{_xSeriesTopPath}/c:strLit", _ns);
				xmlNode4?.ParentNode.RemoveChild(xmlNode4);
			}
		}
	}

	public ExcelChartTrendlineCollection TrendLines
	{
		get
		{
			if (_trendLines == null)
			{
				_trendLines = new ExcelChartTrendlineCollection(this);
			}
			return _trendLines;
		}
	}

	public ExcelDrawingFill Fill
	{
		get
		{
			if (_fill == null)
			{
				_fill = new ExcelDrawingFill(base.NameSpaceManager, base.TopNode, "c:spPr");
			}
			return _fill;
		}
	}

	public ExcelDrawingBorder Border
	{
		get
		{
			if (_border == null)
			{
				_border = new ExcelDrawingBorder(base.NameSpaceManager, base.TopNode, "c:spPr/a:ln");
			}
			return _border;
		}
	}

	internal ExcelChartSerie(ExcelChartSeries chartSeries, XmlNamespaceManager ns, XmlNode node, bool isPivot)
		: base(ns, node)
	{
		_chartSeries = chartSeries;
		_node = node;
		_ns = ns;
		base.SchemaNodeOrder = new string[16]
		{
			"idx", "order", "spPr", "tx", "marker", "trendline", "explosion", "invertIfNegative", "dLbls", "cat",
			"val", "xVal", "yVal", "bubbleSize", "bubble3D", "smooth"
		};
		if (chartSeries.Chart.ChartType == eChartType.XYScatter || chartSeries.Chart.ChartType == eChartType.XYScatterLines || chartSeries.Chart.ChartType == eChartType.XYScatterLinesNoMarkers || chartSeries.Chart.ChartType == eChartType.XYScatterSmooth || chartSeries.Chart.ChartType == eChartType.XYScatterSmoothNoMarkers || chartSeries.Chart.ChartType == eChartType.Bubble || chartSeries.Chart.ChartType == eChartType.Bubble3DEffect)
		{
			_seriesTopPath = "c:yVal";
			_xSeriesTopPath = "c:xVal";
		}
		else
		{
			_seriesTopPath = "c:val";
			_xSeriesTopPath = "c:cat";
		}
		_seriesPath = string.Format(_seriesPath, _seriesTopPath);
		string xSeriesPath = string.Format(_xSeriesPath, _xSeriesTopPath, isPivot ? "c:multiLvlStrRef" : "c:numRef");
		string text = string.Format(_xSeriesPath, _xSeriesTopPath, isPivot ? "c:multiLvlStrRef" : "c:strRef");
		if (ExistNode(text))
		{
			_xSeriesPath = text;
		}
		else
		{
			_xSeriesPath = xSeriesPath;
		}
	}

	internal void SetID(string id)
	{
		SetXmlNodeString("c:idx/@val", id);
		SetXmlNodeString("c:order/@val", id);
	}

	private void Cleartx()
	{
		XmlNode xmlNode = base.TopNode.SelectSingleNode("c:tx", base.NameSpaceManager);
		if (xmlNode != null)
		{
			xmlNode.InnerXml = "";
		}
	}
}
