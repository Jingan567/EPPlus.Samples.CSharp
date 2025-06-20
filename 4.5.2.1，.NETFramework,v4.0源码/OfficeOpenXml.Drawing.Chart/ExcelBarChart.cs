using System;
using System.Globalization;
using System.Xml;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Table.PivotTable;

namespace OfficeOpenXml.Drawing.Chart;

public sealed class ExcelBarChart : ExcelChart
{
	private string _directionPath = "c:barDir/@val";

	private string _shapePath = "c:shape/@val";

	private ExcelChartDataLabel _DataLabel;

	private string _gapWidthPath = "c:gapWidth/@val";

	public eDirection Direction
	{
		get
		{
			return GetDirectionEnum(_chartXmlHelper.GetXmlNodeString(_directionPath));
		}
		internal set
		{
			_chartXmlHelper.SetXmlNodeString(_directionPath, GetDirectionText(value));
		}
	}

	public eShape Shape
	{
		get
		{
			return GetShapeEnum(_chartXmlHelper.GetXmlNodeString(_shapePath));
		}
		internal set
		{
			_chartXmlHelper.SetXmlNodeString(_shapePath, GetShapeText(value));
		}
	}

	public ExcelChartDataLabel DataLabel
	{
		get
		{
			if (_DataLabel == null)
			{
				_DataLabel = new ExcelChartDataLabel(base.NameSpaceManager, base.ChartNode);
			}
			return _DataLabel;
		}
	}

	public int GapWidth
	{
		get
		{
			return _chartXmlHelper.GetXmlNodeInt(_gapWidthPath);
		}
		set
		{
			_chartXmlHelper.SetXmlNodeString(_gapWidthPath, value.ToString(CultureInfo.InvariantCulture));
		}
	}

	internal ExcelBarChart(ExcelDrawings drawings, XmlNode node, eChartType type, ExcelChart topChart, ExcelPivotTable PivotTableSource)
		: base(drawings, node, type, topChart, PivotTableSource)
	{
		SetChartNodeText("");
		SetTypeProperties(drawings, type);
	}

	internal ExcelBarChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode)
		: base(drawings, node, uriChart, part, chartXml, chartNode)
	{
		SetChartNodeText(chartNode.Name);
	}

	internal ExcelBarChart(ExcelChart topChart, XmlNode chartNode)
		: base(topChart, chartNode)
	{
		SetChartNodeText(chartNode.Name);
	}

	private void SetChartNodeText(string chartNodeText)
	{
		if (string.IsNullOrEmpty(chartNodeText))
		{
			chartNodeText = GetChartNodeText();
		}
	}

	private void SetTypeProperties(ExcelDrawings drawings, eChartType type)
	{
		switch (type)
		{
		case eChartType.BarClustered:
		case eChartType.BarStacked:
		case eChartType.BarStacked100:
		case eChartType.BarClustered3D:
		case eChartType.BarStacked3D:
		case eChartType.BarStacked1003D:
		case eChartType.CylinderBarClustered:
		case eChartType.CylinderBarStacked:
		case eChartType.CylinderBarStacked100:
		case eChartType.ConeBarClustered:
		case eChartType.ConeBarStacked:
		case eChartType.ConeBarStacked100:
		case eChartType.PyramidBarClustered:
		case eChartType.PyramidBarStacked:
		case eChartType.PyramidBarStacked100:
			Direction = eDirection.Bar;
			break;
		case eChartType.Column3D:
		case eChartType.ColumnClustered:
		case eChartType.ColumnStacked:
		case eChartType.ColumnStacked100:
		case eChartType.ColumnClustered3D:
		case eChartType.ColumnStacked3D:
		case eChartType.ColumnStacked1003D:
		case eChartType.CylinderColClustered:
		case eChartType.CylinderColStacked:
		case eChartType.CylinderColStacked100:
		case eChartType.CylinderCol:
		case eChartType.ConeColClustered:
		case eChartType.ConeColStacked:
		case eChartType.ConeColStacked100:
		case eChartType.ConeCol:
		case eChartType.PyramidColClustered:
		case eChartType.PyramidColStacked:
		case eChartType.PyramidColStacked100:
		case eChartType.PyramidCol:
			Direction = eDirection.Column;
			break;
		}
		switch (type)
		{
		case eChartType.Column3D:
		case eChartType.ColumnClustered3D:
		case eChartType.ColumnStacked3D:
		case eChartType.ColumnStacked1003D:
		case eChartType.BarClustered3D:
		case eChartType.BarStacked3D:
		case eChartType.BarStacked1003D:
			Shape = eShape.Box;
			break;
		case eChartType.CylinderColClustered:
		case eChartType.CylinderColStacked:
		case eChartType.CylinderColStacked100:
		case eChartType.CylinderBarClustered:
		case eChartType.CylinderBarStacked:
		case eChartType.CylinderBarStacked100:
		case eChartType.CylinderCol:
			Shape = eShape.Cylinder;
			break;
		case eChartType.ConeColClustered:
		case eChartType.ConeColStacked:
		case eChartType.ConeColStacked100:
		case eChartType.ConeBarClustered:
		case eChartType.ConeBarStacked:
		case eChartType.ConeBarStacked100:
		case eChartType.ConeCol:
			Shape = eShape.Cone;
			break;
		case eChartType.PyramidColClustered:
		case eChartType.PyramidColStacked:
		case eChartType.PyramidColStacked100:
		case eChartType.PyramidBarClustered:
		case eChartType.PyramidBarStacked:
		case eChartType.PyramidBarStacked100:
		case eChartType.PyramidCol:
			Shape = eShape.Pyramid;
			break;
		}
	}

	private string GetDirectionText(eDirection direction)
	{
		if (direction == eDirection.Bar)
		{
			return "bar";
		}
		return "col";
	}

	private eDirection GetDirectionEnum(string direction)
	{
		if (direction == "bar")
		{
			return eDirection.Bar;
		}
		return eDirection.Column;
	}

	private string GetShapeText(eShape Shape)
	{
		return Shape switch
		{
			eShape.Box => "box", 
			eShape.Cone => "cone", 
			eShape.ConeToMax => "coneToMax", 
			eShape.Cylinder => "cylinder", 
			eShape.Pyramid => "pyramid", 
			eShape.PyramidToMax => "pyramidToMax", 
			_ => "box", 
		};
	}

	private eShape GetShapeEnum(string text)
	{
		return text switch
		{
			"box" => eShape.Box, 
			"cone" => eShape.Cone, 
			"coneToMax" => eShape.ConeToMax, 
			"cylinder" => eShape.Cylinder, 
			"pyramid" => eShape.Pyramid, 
			"pyramidToMax" => eShape.PyramidToMax, 
			_ => eShape.Box, 
		};
	}

	internal override eChartType GetChartType(string name)
	{
		if (name == "barChart")
		{
			if (Direction == eDirection.Bar)
			{
				if (base.Grouping == eGrouping.Stacked)
				{
					return eChartType.BarStacked;
				}
				if (base.Grouping == eGrouping.PercentStacked)
				{
					return eChartType.BarStacked100;
				}
				return eChartType.BarClustered;
			}
			if (base.Grouping == eGrouping.Stacked)
			{
				return eChartType.ColumnStacked;
			}
			if (base.Grouping == eGrouping.PercentStacked)
			{
				return eChartType.ColumnStacked100;
			}
			return eChartType.ColumnClustered;
		}
		if (name == "bar3DChart")
		{
			if (Shape == eShape.Box)
			{
				if (Direction == eDirection.Bar)
				{
					if (base.Grouping == eGrouping.Stacked)
					{
						return eChartType.BarStacked3D;
					}
					if (base.Grouping == eGrouping.PercentStacked)
					{
						return eChartType.BarStacked1003D;
					}
					return eChartType.BarClustered3D;
				}
				if (base.Grouping == eGrouping.Stacked)
				{
					return eChartType.ColumnStacked3D;
				}
				if (base.Grouping == eGrouping.PercentStacked)
				{
					return eChartType.ColumnStacked1003D;
				}
				return eChartType.ColumnClustered3D;
			}
			if (Shape == eShape.Cone || Shape == eShape.ConeToMax)
			{
				if (Direction != eDirection.Bar)
				{
					if (base.Grouping == eGrouping.Stacked)
					{
						return eChartType.ConeColStacked;
					}
					if (base.Grouping == eGrouping.PercentStacked)
					{
						return eChartType.ConeColStacked100;
					}
					if (base.Grouping == eGrouping.Clustered)
					{
						return eChartType.ConeColClustered;
					}
					return eChartType.ConeCol;
				}
				if (base.Grouping == eGrouping.Stacked)
				{
					return eChartType.ConeBarStacked;
				}
				if (base.Grouping == eGrouping.PercentStacked)
				{
					return eChartType.ConeBarStacked100;
				}
				if (base.Grouping == eGrouping.Clustered)
				{
					return eChartType.ConeBarClustered;
				}
			}
			if (Shape == eShape.Cylinder)
			{
				if (Direction != eDirection.Bar)
				{
					if (base.Grouping == eGrouping.Stacked)
					{
						return eChartType.CylinderColStacked;
					}
					if (base.Grouping == eGrouping.PercentStacked)
					{
						return eChartType.CylinderColStacked100;
					}
					if (base.Grouping == eGrouping.Clustered)
					{
						return eChartType.CylinderColClustered;
					}
					return eChartType.CylinderCol;
				}
				if (base.Grouping == eGrouping.Stacked)
				{
					return eChartType.CylinderBarStacked;
				}
				if (base.Grouping == eGrouping.PercentStacked)
				{
					return eChartType.CylinderBarStacked100;
				}
				if (base.Grouping == eGrouping.Clustered)
				{
					return eChartType.CylinderBarClustered;
				}
			}
			if (Shape == eShape.Pyramid || Shape == eShape.PyramidToMax)
			{
				if (Direction != eDirection.Bar)
				{
					if (base.Grouping == eGrouping.Stacked)
					{
						return eChartType.PyramidColStacked;
					}
					if (base.Grouping == eGrouping.PercentStacked)
					{
						return eChartType.PyramidColStacked100;
					}
					if (base.Grouping == eGrouping.Clustered)
					{
						return eChartType.PyramidColClustered;
					}
					return eChartType.PyramidCol;
				}
				if (base.Grouping == eGrouping.Stacked)
				{
					return eChartType.PyramidBarStacked;
				}
				if (base.Grouping == eGrouping.PercentStacked)
				{
					return eChartType.PyramidBarStacked100;
				}
				if (base.Grouping == eGrouping.Clustered)
				{
					return eChartType.PyramidBarClustered;
				}
			}
		}
		return base.GetChartType(name);
	}
}
