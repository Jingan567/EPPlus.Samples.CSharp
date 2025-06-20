using System;
using System.Collections;
using System.Collections.Generic;

namespace OfficeOpenXml.Drawing.Chart;

public class ExcelChartCollection : IEnumerable<ExcelChart>, IEnumerable
{
	private List<ExcelChart> _list = new List<ExcelChart>();

	private ExcelChart _topChart;

	public int Count => _list.Count;

	public ExcelChart this[int PositionID] => _list[PositionID];

	internal ExcelChartCollection(ExcelChart chart)
	{
		_topChart = chart;
		_list.Add(chart);
	}

	internal void Add(ExcelChart chart)
	{
		_list.Add(chart);
	}

	public ExcelChart Add(eChartType chartType)
	{
		if (_topChart.PivotTableSource != null)
		{
			throw new InvalidOperationException("Can not add other charttypes to a pivot chart");
		}
		if (ExcelChart.IsType3D(chartType) || _list[0].IsType3D())
		{
			throw new InvalidOperationException("3D charts can not be combined with other charttypes");
		}
		_ = _list[_list.Count - 1].TopNode;
		ExcelChart newChart = ExcelChart.GetNewChart(_topChart.WorkSheet.Drawings, _topChart.TopNode, chartType, _topChart, null);
		_list.Add(newChart);
		return newChart;
	}

	IEnumerator<ExcelChart> IEnumerable<ExcelChart>.GetEnumerator()
	{
		return _list.GetEnumerator();
	}

	IEnumerator IEnumerable.GetEnumerator()
	{
		return _list.GetEnumerator();
	}
}
