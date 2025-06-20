using System;
using System.Collections;
using System.Collections.Generic;
using System.Xml;
using OfficeOpenXml.Table.PivotTable;

namespace OfficeOpenXml.Drawing.Chart;

public class ExcelChartSeries : XmlHelper, IEnumerable
{
	private List<ExcelChartSerie> _list = new List<ExcelChartSerie>();

	internal ExcelChart _chart;

	private XmlNode _node;

	private XmlNamespaceManager _ns;

	private bool _isPivot;

	public ExcelChartSerie this[int PositionID] => _list[PositionID];

	public int Count => _list.Count;

	public ExcelChart Chart => _chart;

	internal ExcelChartSeries(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, bool isPivot)
		: base(ns, node)
	{
		_ns = ns;
		_chart = chart;
		_node = node;
		_isPivot = isPivot;
		base.SchemaNodeOrder = new string[14]
		{
			"view3D", "plotArea", "barDir", "grouping", "scatterStyle", "varyColors", "ser", "explosion", "dLbls", "firstSliceAng",
			"holeSize", "shape", "legend", "axId"
		};
		foreach (XmlNode item2 in node.SelectNodes("c:ser", ns))
		{
			ExcelChartSerie item = ((!(chart.ChartNode.LocalName == "scatterChart")) ? ((!(chart.ChartNode.LocalName == "lineChart")) ? ((!(chart.ChartNode.LocalName == "pieChart") && !(chart.ChartNode.LocalName == "ofPieChart") && !(chart.ChartNode.LocalName == "pie3DChart") && !(chart.ChartNode.LocalName == "doughnutChart")) ? ((!(chart.ChartNode.LocalName == "bubbleChart")) ? new ExcelChartSerie(this, ns, item2, isPivot) : new ExcelBubbleChartSerie(this, ns, item2, isPivot)) : new ExcelPieChartSerie(this, ns, item2, isPivot)) : new ExcelLineChartSerie(this, ns, item2, isPivot)) : new ExcelScatterChartSerie(this, ns, item2, isPivot));
			_list.Add(item);
		}
	}

	public IEnumerator GetEnumerator()
	{
		return _list.GetEnumerator();
	}

	public void Delete(int PositionID)
	{
		ExcelChartSerie excelChartSerie = _list[PositionID];
		excelChartSerie.TopNode.ParentNode.RemoveChild(excelChartSerie.TopNode);
		_list.RemoveAt(PositionID);
	}

	public virtual ExcelChartSerie Add(ExcelRangeBase Serie, ExcelRangeBase XSerie)
	{
		if (_chart.PivotTableSource != null)
		{
			throw new InvalidOperationException("Can't add a serie to a pivotchart");
		}
		return AddSeries(Serie.FullAddressAbsolute, XSerie.FullAddressAbsolute, "");
	}

	public virtual ExcelChartSerie Add(string SerieAddress, string XSerieAddress)
	{
		if (_chart.PivotTableSource != null)
		{
			throw new InvalidOperationException("Can't add a serie to a pivotchart");
		}
		return AddSeries(SerieAddress, XSerieAddress, "");
	}

	protected internal ExcelChartSerie AddSeries(string SeriesAddress, string XSeriesAddress, string bubbleSizeAddress)
	{
		XmlElement xmlElement = _node.OwnerDocument.CreateElement("ser", "http://schemas.openxmlformats.org/drawingml/2006/chart");
		XmlNodeList xmlNodeList = _node.SelectNodes("c:ser", _ns);
		if (xmlNodeList.Count > 0)
		{
			_node.InsertAfter(xmlElement, xmlNodeList[xmlNodeList.Count - 1]);
		}
		else
		{
			InserAfter(_node, "c:varyColors,c:grouping,c:barDir,c:scatterStyle,c:ofPieType", xmlElement);
		}
		int num = FindIndex();
		xmlElement.InnerXml = string.Format("<c:idx val=\"{1}\" /><c:order val=\"{1}\" /><c:tx><c:strRef><c:f></c:f><c:strCache><c:ptCount val=\"1\" /></c:strCache></c:strRef></c:tx>{5}{0}{2}{3}{4}", AddExplosion(Chart.ChartType), num, AddScatterPoint(Chart.ChartType), AddAxisNodes(Chart.ChartType), AddSmooth(Chart.ChartType), AddMarker(Chart.ChartType));
		ExcelChartSerie excelChartSerie;
		switch (Chart.ChartType)
		{
		case eChartType.Bubble:
		case eChartType.Bubble3DEffect:
			excelChartSerie = new ExcelBubbleChartSerie(this, base.NameSpaceManager, xmlElement, _isPivot)
			{
				Bubble3D = (Chart.ChartType == eChartType.Bubble3DEffect),
				Series = SeriesAddress,
				XSeries = XSeriesAddress,
				BubbleSize = bubbleSizeAddress
			};
			break;
		case eChartType.XYScatter:
		case eChartType.XYScatterSmooth:
		case eChartType.XYScatterSmoothNoMarkers:
		case eChartType.XYScatterLines:
		case eChartType.XYScatterLinesNoMarkers:
			excelChartSerie = new ExcelScatterChartSerie(this, base.NameSpaceManager, xmlElement, _isPivot);
			break;
		case eChartType.Radar:
		case eChartType.RadarMarkers:
		case eChartType.RadarFilled:
			excelChartSerie = new ExcelRadarChartSerie(this, base.NameSpaceManager, xmlElement, _isPivot);
			break;
		case eChartType.Surface:
		case eChartType.SurfaceWireframe:
		case eChartType.SurfaceTopView:
		case eChartType.SurfaceTopViewWireframe:
			excelChartSerie = new ExcelSurfaceChartSerie(this, base.NameSpaceManager, xmlElement, _isPivot);
			break;
		case eChartType.Doughnut:
		case eChartType.Pie3D:
		case eChartType.Pie:
		case eChartType.PieOfPie:
		case eChartType.PieExploded:
		case eChartType.PieExploded3D:
		case eChartType.BarOfPie:
		case eChartType.DoughnutExploded:
			excelChartSerie = new ExcelPieChartSerie(this, base.NameSpaceManager, xmlElement, _isPivot);
			break;
		case eChartType.Line:
		case eChartType.LineStacked:
		case eChartType.LineStacked100:
		case eChartType.LineMarkers:
		case eChartType.LineMarkersStacked:
		case eChartType.LineMarkersStacked100:
			excelChartSerie = new ExcelLineChartSerie(this, base.NameSpaceManager, xmlElement, _isPivot);
			if (Chart.ChartType == eChartType.LineMarkers || Chart.ChartType == eChartType.LineMarkersStacked || Chart.ChartType == eChartType.LineMarkersStacked100)
			{
				((ExcelLineChartSerie)excelChartSerie).Marker = eMarkerStyle.Square;
			}
			((ExcelLineChartSerie)excelChartSerie).Smooth = ((ExcelLineChart)Chart).Smooth;
			break;
		case eChartType.ColumnClustered:
		case eChartType.ColumnStacked:
		case eChartType.ColumnStacked100:
		case eChartType.ColumnClustered3D:
		case eChartType.ColumnStacked3D:
		case eChartType.ColumnStacked1003D:
		case eChartType.BarClustered:
		case eChartType.BarStacked:
		case eChartType.BarStacked100:
		case eChartType.BarClustered3D:
		case eChartType.BarStacked3D:
		case eChartType.BarStacked1003D:
		case eChartType.CylinderColClustered:
		case eChartType.CylinderColStacked:
		case eChartType.CylinderColStacked100:
		case eChartType.CylinderBarClustered:
		case eChartType.CylinderBarStacked:
		case eChartType.CylinderBarStacked100:
		case eChartType.CylinderCol:
		case eChartType.ConeColClustered:
		case eChartType.ConeColStacked:
		case eChartType.ConeColStacked100:
		case eChartType.ConeBarClustered:
		case eChartType.ConeBarStacked:
		case eChartType.ConeBarStacked100:
		case eChartType.ConeCol:
		case eChartType.PyramidColClustered:
		case eChartType.PyramidColStacked:
		case eChartType.PyramidColStacked100:
		case eChartType.PyramidBarClustered:
		case eChartType.PyramidBarStacked:
		case eChartType.PyramidBarStacked100:
		case eChartType.PyramidCol:
			excelChartSerie = new ExcelBarChartSerie(this, base.NameSpaceManager, xmlElement, _isPivot);
			((ExcelBarChartSerie)excelChartSerie).InvertIfNegative = false;
			break;
		default:
			excelChartSerie = new ExcelChartSerie(this, base.NameSpaceManager, xmlElement, _isPivot);
			break;
		}
		excelChartSerie.Series = SeriesAddress;
		excelChartSerie.XSeries = XSeriesAddress;
		_list.Add(excelChartSerie);
		return excelChartSerie;
	}

	internal void AddPivotSerie(ExcelPivotTable pivotTableSource)
	{
		ExcelRange excelRange = pivotTableSource.WorkSheet.Cells[pivotTableSource.Address.Address];
		_isPivot = true;
		AddSeries(excelRange.Offset(0, 1, excelRange._toRow - excelRange._fromRow + 1, 1).FullAddressAbsolute, excelRange.Offset(0, 0, excelRange._toRow - excelRange._fromRow + 1, 1).FullAddressAbsolute, "");
	}

	private int FindIndex()
	{
		int num = 0;
		int num2 = 0;
		if (_chart.PlotArea.ChartTypes.Count > 1)
		{
			foreach (ExcelChart item in (IEnumerable<ExcelChart>)_chart.PlotArea.ChartTypes)
			{
				if (num2 > 0)
				{
					foreach (ExcelChartSerie item2 in item.Series)
					{
						int num3 = num2 + 1;
						num2 = num3;
						item2.SetID(num3.ToString());
					}
				}
				else if (item == _chart)
				{
					num += _list.Count + 1;
					num2 = num;
				}
				else
				{
					num += item.Series.Count;
				}
			}
			return num - 1;
		}
		return _list.Count;
	}

	private string AddMarker(eChartType chartType)
	{
		if (chartType == eChartType.Line || chartType == eChartType.LineStacked || chartType == eChartType.LineStacked100 || chartType == eChartType.XYScatterLines || chartType == eChartType.XYScatterSmooth || chartType == eChartType.XYScatterLinesNoMarkers || chartType == eChartType.XYScatterSmoothNoMarkers)
		{
			return "<c:marker><c:symbol val=\"none\" /></c:marker>";
		}
		return "";
	}

	private string AddScatterPoint(eChartType chartType)
	{
		if (chartType == eChartType.XYScatter)
		{
			return "<c:spPr><a:ln w=\"28575\"><a:noFill /></a:ln></c:spPr>";
		}
		return "";
	}

	private string AddAxisNodes(eChartType chartType)
	{
		if (chartType == eChartType.XYScatter || chartType == eChartType.XYScatterLines || chartType == eChartType.XYScatterLinesNoMarkers || chartType == eChartType.XYScatterSmooth || chartType == eChartType.XYScatterSmoothNoMarkers || chartType == eChartType.Bubble || chartType == eChartType.Bubble3DEffect)
		{
			return "<c:xVal /><c:yVal />";
		}
		return "<c:val />";
	}

	private string AddExplosion(eChartType chartType)
	{
		if (chartType == eChartType.PieExploded3D || chartType == eChartType.PieExploded || chartType == eChartType.DoughnutExploded)
		{
			return "<c:explosion val=\"25\" />";
		}
		return "";
	}

	private string AddSmooth(eChartType chartType)
	{
		if (chartType == eChartType.XYScatterSmooth || chartType == eChartType.XYScatterSmoothNoMarkers)
		{
			return "<c:smooth val=\"1\" />";
		}
		return "";
	}
}
