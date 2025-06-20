using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart;

public sealed class ExcelChartPlotArea : XmlHelper
{
	private ExcelChart _firstChart;

	private ExcelChartCollection _chartTypes;

	private ExcelChartDataTable _dataTable;

	private ExcelDrawingFill _fill;

	private ExcelDrawingBorder _border;

	public ExcelChartCollection ChartTypes
	{
		get
		{
			if (_chartTypes == null)
			{
				_chartTypes = new ExcelChartCollection(_firstChart);
			}
			return _chartTypes;
		}
	}

	public ExcelChartDataTable DataTable => _dataTable;

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

	internal ExcelChartPlotArea(XmlNamespaceManager ns, XmlNode node, ExcelChart firstChart)
		: base(ns, node)
	{
		_firstChart = firstChart;
		if (base.TopNode.SelectSingleNode("c:dTable", base.NameSpaceManager) != null)
		{
			_dataTable = new ExcelChartDataTable(base.NameSpaceManager, base.TopNode);
		}
	}

	public ExcelChartDataTable CreateDataTable()
	{
		if (_dataTable != null)
		{
			throw new InvalidOperationException("Data table already exists");
		}
		_dataTable = new ExcelChartDataTable(base.NameSpaceManager, base.TopNode);
		return _dataTable;
	}

	public void RemoveDataTable()
	{
		DeleteAllNode("c:dTable");
		_dataTable = null;
	}
}
