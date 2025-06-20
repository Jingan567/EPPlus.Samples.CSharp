using System;
using System.Xml;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Table.PivotTable;

namespace OfficeOpenXml.Drawing.Chart;

public sealed class ExcelSurfaceChart : ExcelChart
{
	private ExcelChartSurface _floor;

	private ExcelChartSurface _sideWall;

	private ExcelChartSurface _backWall;

	private const string WIREFRAME_PATH = "c:wireframe/@val";

	public ExcelChartSurface Floor => _floor;

	public ExcelChartSurface SideWall => _sideWall;

	public ExcelChartSurface BackWall => _backWall;

	public bool Wireframe
	{
		get
		{
			return _chartXmlHelper.GetXmlNodeBool("c:wireframe/@val");
		}
		set
		{
			_chartXmlHelper.SetXmlNodeBool("c:wireframe/@val", value);
		}
	}

	internal ExcelSurfaceChart(ExcelDrawings drawings, XmlNode node, eChartType type, ExcelChart topChart, ExcelPivotTable PivotTableSource)
		: base(drawings, node, type, topChart, PivotTableSource)
	{
		Init();
	}

	internal ExcelSurfaceChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode)
		: base(drawings, node, uriChart, part, chartXml, chartNode)
	{
		Init();
	}

	internal ExcelSurfaceChart(ExcelChart topChart, XmlNode chartNode)
		: base(topChart, chartNode)
	{
		Init();
	}

	private void Init()
	{
		_floor = new ExcelChartSurface(base.NameSpaceManager, _chartXmlHelper.TopNode.SelectSingleNode("c:floor", base.NameSpaceManager));
		_backWall = new ExcelChartSurface(base.NameSpaceManager, _chartXmlHelper.TopNode.SelectSingleNode("c:sideWall", base.NameSpaceManager));
		_sideWall = new ExcelChartSurface(base.NameSpaceManager, _chartXmlHelper.TopNode.SelectSingleNode("c:backWall", base.NameSpaceManager));
		SetTypeProperties();
	}

	internal void SetTypeProperties()
	{
		if (base.ChartType == eChartType.SurfaceWireframe || base.ChartType == eChartType.SurfaceTopViewWireframe)
		{
			Wireframe = true;
		}
		else
		{
			Wireframe = false;
		}
		if (base.ChartType == eChartType.SurfaceTopView || base.ChartType == eChartType.SurfaceTopViewWireframe)
		{
			base.View3D.RotY = 0m;
			base.View3D.RotX = 90m;
		}
		else
		{
			base.View3D.RotY = 20m;
			base.View3D.RotX = 15m;
		}
		base.View3D.RightAngleAxes = false;
		base.View3D.Perspective = 0m;
		base.Axis[1].CrossBetween = eCrossBetween.MidCat;
	}

	internal override eChartType GetChartType(string name)
	{
		if (Wireframe)
		{
			if (name == "surfaceChart")
			{
				return eChartType.SurfaceTopViewWireframe;
			}
			return eChartType.SurfaceWireframe;
		}
		if (name == "surfaceChart")
		{
			return eChartType.SurfaceTopView;
		}
		return eChartType.Surface;
	}
}
