using System;
using System.Collections;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart;

public class ExcelChartTrendlineCollection : IEnumerable<ExcelChartTrendline>, IEnumerable
{
	private List<ExcelChartTrendline> _list = new List<ExcelChartTrendline>();

	private ExcelChartSerie _serie;

	internal ExcelChartTrendlineCollection(ExcelChartSerie serie)
	{
		_serie = serie;
		foreach (XmlNode item in _serie.TopNode.SelectNodes("c:trendline", _serie.NameSpaceManager))
		{
			_list.Add(new ExcelChartTrendline(_serie.NameSpaceManager, item));
		}
	}

	public ExcelChartTrendline Add(eTrendLine Type)
	{
		if (_serie._chartSeries._chart.IsType3D() || _serie._chartSeries._chart.IsTypePercentStacked() || _serie._chartSeries._chart.IsTypeStacked() || _serie._chartSeries._chart.IsTypePieDoughnut())
		{
			throw new ArgumentException("Trendlines don't apply to 3d-charts, stacked charts, pie charts or doughnut charts");
		}
		XmlNode xmlNode;
		if (_list.Count > 0)
		{
			xmlNode = _list[_list.Count - 1].TopNode;
		}
		else
		{
			xmlNode = _serie.TopNode.SelectSingleNode("c:marker", _serie.NameSpaceManager);
			if (xmlNode == null)
			{
				xmlNode = _serie.TopNode.SelectSingleNode("c:tx", _serie.NameSpaceManager);
				if (xmlNode == null)
				{
					xmlNode = _serie.TopNode.SelectSingleNode("c:order", _serie.NameSpaceManager);
				}
			}
		}
		XmlElement xmlElement = _serie.TopNode.OwnerDocument.CreateElement("c", "trendline", "http://schemas.openxmlformats.org/drawingml/2006/chart");
		_serie.TopNode.InsertAfter(xmlElement, xmlNode);
		return new ExcelChartTrendline(_serie.NameSpaceManager, xmlElement)
		{
			Type = Type
		};
	}

	IEnumerator<ExcelChartTrendline> IEnumerable<ExcelChartTrendline>.GetEnumerator()
	{
		return _list.GetEnumerator();
	}

	IEnumerator IEnumerable.GetEnumerator()
	{
		return _list.GetEnumerator();
	}
}
