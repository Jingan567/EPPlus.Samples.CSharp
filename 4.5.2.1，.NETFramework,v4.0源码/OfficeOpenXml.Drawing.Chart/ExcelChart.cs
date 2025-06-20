using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Drawing.Chart;

public class ExcelChart : ExcelDrawing
{
	private const string rootPath = "c:chartSpace/c:chart/c:plotArea";

	protected internal ExcelChartSeries _chartSeries;

	internal ExcelChartAxis[] _axis;

	protected XmlHelper _chartXmlHelper;

	protected internal XmlNode _chartNode;

	private bool _secondaryAxis;

	private const string _roundedCornersPath = "../../../c:roundedCorners/@val";

	private const string _plotVisibleOnlyPath = "../../c:plotVisOnly/@val";

	private const string _displayBlanksAsPath = "../../c:dispBlanksAs/@val";

	private const string _showDLblsOverMax = "../../c:showDLblsOverMax/@val";

	private ExcelChartPlotArea _plotArea;

	private ExcelChartLegend _legend;

	private ExcelDrawingBorder _border;

	private ExcelDrawingFill _fill;

	private string _groupingPath = "c:grouping/@val";

	private string _varyColorsPath = "c:varyColors/@val";

	private ExcelChartTitle _title;

	public ExcelWorksheet WorkSheet { get; internal set; }

	public XmlDocument ChartXml { get; internal set; }

	public eChartType ChartType { get; internal set; }

	internal XmlNode ChartNode => _chartNode;

	public ExcelChartTitle Title
	{
		get
		{
			if (_title == null)
			{
				_title = new ExcelChartTitle(base.NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace/c:chart", base.NameSpaceManager));
			}
			return _title;
		}
	}

	public virtual ExcelChartSeries Series => _chartSeries;

	public ExcelChartAxis[] Axis => _axis;

	public ExcelChartAxis XAxis { get; private set; }

	public ExcelChartAxis YAxis { get; private set; }

	public bool UseSecondaryAxis
	{
		get
		{
			return _secondaryAxis;
		}
		set
		{
			if (_secondaryAxis == value)
			{
				return;
			}
			if (value)
			{
				if (IsTypePieDoughnut())
				{
					throw new Exception("Pie charts do not support axis");
				}
				if (!HasPrimaryAxis())
				{
					throw new Exception("Can't set to secondary axis when no serie uses the primary axis");
				}
				if (Axis.Length == 2)
				{
					AddAxis();
				}
				XmlNodeList xmlNodeList = ChartNode.SelectNodes("c:axId", base.NameSpaceManager);
				xmlNodeList[0].Attributes["val"].Value = Axis[2].Id;
				xmlNodeList[1].Attributes["val"].Value = Axis[3].Id;
				XAxis = Axis[2];
				YAxis = Axis[3];
			}
			else
			{
				XmlNodeList xmlNodeList2 = ChartNode.SelectNodes("c:axId", base.NameSpaceManager);
				xmlNodeList2[0].Attributes["val"].Value = Axis[0].Id;
				xmlNodeList2[1].Attributes["val"].Value = Axis[1].Id;
				XAxis = Axis[0];
				YAxis = Axis[1];
			}
			_secondaryAxis = value;
		}
	}

	public eChartStyle Style
	{
		get
		{
			XmlNode xmlNode = ChartXml.SelectSingleNode("c:chartSpace/c:style/@val", base.NameSpaceManager);
			if (xmlNode == null)
			{
				return eChartStyle.None;
			}
			if (int.TryParse(xmlNode.Value, NumberStyles.Number, CultureInfo.InvariantCulture, out var result))
			{
				return (eChartStyle)result;
			}
			return eChartStyle.None;
		}
		set
		{
			if (value == eChartStyle.None)
			{
				if (ChartXml.SelectSingleNode("c:chartSpace/c:style", base.NameSpaceManager) is XmlElement xmlElement)
				{
					xmlElement.ParentNode.RemoveChild(xmlElement);
				}
			}
			else
			{
				XmlElement xmlElement2 = ChartXml.CreateElement("c:style", "http://schemas.openxmlformats.org/drawingml/2006/chart");
				int num = (int)value;
				xmlElement2.SetAttribute("val", num.ToString());
				XmlElement xmlElement3 = ChartXml.SelectSingleNode("c:chartSpace", base.NameSpaceManager) as XmlElement;
				xmlElement3.InsertBefore(xmlElement2, xmlElement3.SelectSingleNode("c:chart", base.NameSpaceManager));
			}
		}
	}

	public bool RoundedCorners
	{
		get
		{
			return _chartXmlHelper.GetXmlNodeBool("../../../c:roundedCorners/@val");
		}
		set
		{
			_chartXmlHelper.SetXmlNodeBool("../../../c:roundedCorners/@val", value);
		}
	}

	public bool ShowHiddenData
	{
		get
		{
			return !_chartXmlHelper.GetXmlNodeBool("../../c:plotVisOnly/@val");
		}
		set
		{
			_chartXmlHelper.SetXmlNodeBool("../../c:plotVisOnly/@val", !value);
		}
	}

	public eDisplayBlanksAs DisplayBlanksAs
	{
		get
		{
			string xmlNodeString = _chartXmlHelper.GetXmlNodeString("../../c:dispBlanksAs/@val");
			if (string.IsNullOrEmpty(xmlNodeString))
			{
				return eDisplayBlanksAs.Zero;
			}
			return (eDisplayBlanksAs)Enum.Parse(typeof(eDisplayBlanksAs), xmlNodeString, ignoreCase: true);
		}
		set
		{
			_chartSeries.SetXmlNodeString("../../c:dispBlanksAs/@val", value.ToString().ToLower(CultureInfo.InvariantCulture));
		}
	}

	public bool ShowDataLabelsOverMaximum
	{
		get
		{
			return _chartXmlHelper.GetXmlNodeBool("../../c:showDLblsOverMax/@val", blankValue: true);
		}
		set
		{
			_chartXmlHelper.SetXmlNodeBool("../../c:showDLblsOverMax/@val", value, removeIf: true);
		}
	}

	public ExcelChartPlotArea PlotArea
	{
		get
		{
			if (_plotArea == null)
			{
				_plotArea = new ExcelChartPlotArea(base.NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace/c:chart/c:plotArea", base.NameSpaceManager), this);
			}
			return _plotArea;
		}
	}

	public ExcelChartLegend Legend
	{
		get
		{
			if (_legend == null)
			{
				_legend = new ExcelChartLegend(base.NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace/c:chart/c:legend", base.NameSpaceManager), this);
			}
			return _legend;
		}
	}

	public ExcelDrawingBorder Border
	{
		get
		{
			if (_border == null)
			{
				_border = new ExcelDrawingBorder(base.NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace", base.NameSpaceManager), "c:spPr/a:ln");
			}
			return _border;
		}
	}

	public ExcelDrawingFill Fill
	{
		get
		{
			if (_fill == null)
			{
				_fill = new ExcelDrawingFill(base.NameSpaceManager, ChartXml.SelectSingleNode("c:chartSpace", base.NameSpaceManager), "c:spPr");
			}
			return _fill;
		}
	}

	public ExcelView3D View3D
	{
		get
		{
			if (IsType3D())
			{
				return new ExcelView3D(base.NameSpaceManager, ChartXml.SelectSingleNode("//c:view3D", base.NameSpaceManager));
			}
			throw new Exception("Charttype does not support 3D");
		}
	}

	public eGrouping Grouping
	{
		get
		{
			return GetGroupingEnum(_chartXmlHelper.GetXmlNodeString(_groupingPath));
		}
		internal set
		{
			_chartXmlHelper.SetXmlNodeString(_groupingPath, GetGroupingText(value));
		}
	}

	public bool VaryColors
	{
		get
		{
			return _chartXmlHelper.GetXmlNodeBool(_varyColorsPath);
		}
		set
		{
			if (value)
			{
				_chartXmlHelper.SetXmlNodeString(_varyColorsPath, "1");
			}
			else
			{
				_chartXmlHelper.SetXmlNodeString(_varyColorsPath, "0");
			}
		}
	}

	internal ZipPackagePart Part { get; set; }

	internal Uri UriChart { get; set; }

	internal new string Id => "";

	public ExcelPivotTable PivotTableSource { get; private set; }

	internal ExcelChart(ExcelDrawings drawings, XmlNode node, eChartType type, bool isPivot)
		: base(drawings, node, "xdr:graphicFrame/xdr:nvGraphicFramePr/xdr:cNvPr/@name")
	{
		ChartType = type;
		CreateNewChart(drawings, type, null);
		Init(drawings, _chartNode);
		_chartSeries = new ExcelChartSeries(this, drawings.NameSpaceManager, _chartNode, isPivot);
		SetTypeProperties();
		LoadAxis();
	}

	internal ExcelChart(ExcelDrawings drawings, XmlNode node, eChartType type, ExcelChart topChart, ExcelPivotTable PivotTableSource)
		: base(drawings, node, "xdr:graphicFrame/xdr:nvGraphicFramePr/xdr:cNvPr/@name")
	{
		ChartType = type;
		CreateNewChart(drawings, type, topChart);
		Init(drawings, _chartNode);
		_chartSeries = new ExcelChartSeries(this, drawings.NameSpaceManager, _chartNode, PivotTableSource != null);
		if (PivotTableSource != null)
		{
			SetPivotSource(PivotTableSource);
		}
		SetTypeProperties();
		if (topChart == null)
		{
			LoadAxis();
			return;
		}
		_axis = topChart.Axis;
		if (_axis.Length != 0)
		{
			XAxis = _axis[0];
			YAxis = _axis[1];
		}
	}

	internal ExcelChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode)
		: base(drawings, node, "xdr:graphicFrame/xdr:nvGraphicFramePr/xdr:cNvPr/@name")
	{
		UriChart = uriChart;
		Part = part;
		ChartXml = chartXml;
		_chartNode = chartNode;
		InitChartLoad(drawings, chartNode);
		ChartType = GetChartType(chartNode.LocalName);
	}

	internal ExcelChart(ExcelChart topChart, XmlNode chartNode)
		: base(topChart._drawings, topChart.TopNode, "xdr:graphicFrame/xdr:nvGraphicFramePr/xdr:cNvPr/@name")
	{
		UriChart = topChart.UriChart;
		Part = topChart.Part;
		ChartXml = topChart.ChartXml;
		_plotArea = topChart.PlotArea;
		_chartNode = chartNode;
		InitChartLoad(topChart._drawings, chartNode);
	}

	private void InitChartLoad(ExcelDrawings drawings, XmlNode chartNode)
	{
		bool isPivot = false;
		Init(drawings, chartNode);
		_chartSeries = new ExcelChartSeries(this, drawings.NameSpaceManager, _chartNode, isPivot);
		LoadAxis();
	}

	private void Init(ExcelDrawings drawings, XmlNode chartNode)
	{
		_chartXmlHelper = XmlHelperFactory.Create(drawings.NameSpaceManager, chartNode);
		_chartXmlHelper.SchemaNodeOrder = new string[34]
		{
			"ofPieType", "title", "pivotFmt", "autoTitleDeleted", "view3D", "floor", "sideWall", "backWall", "plotArea", "wireframe",
			"barDir", "grouping", "scatterStyle", "radarStyle", "varyColors", "ser", "dLbls", "bubbleScale", "showNegBubbles", "dropLines",
			"upDownBars", "marker", "smooth", "shape", "legend", "plotVisOnly", "dispBlanksAs", "gapWidth", "showDLblsOverMax", "overlap",
			"bandFmts", "axId", "spPr", "printSettings"
		};
		WorkSheet = drawings.Worksheet;
	}

	private void SetTypeProperties()
	{
		if (IsTypeClustered())
		{
			Grouping = eGrouping.Clustered;
		}
		else if (IsTypeStacked())
		{
			Grouping = eGrouping.Stacked;
		}
		else if (IsTypePercentStacked())
		{
			Grouping = eGrouping.PercentStacked;
		}
		if (IsType3D())
		{
			View3D.RotY = 20m;
			View3D.Perspective = 30m;
			if (IsTypePieDoughnut())
			{
				View3D.RotX = 30m;
			}
			else
			{
				View3D.RotX = 15m;
			}
		}
	}

	private void CreateNewChart(ExcelDrawings drawings, eChartType type, ExcelChart topChart)
	{
		if (topChart == null)
		{
			XmlElement xmlElement = base.TopNode.OwnerDocument.CreateElement("graphicFrame", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
			xmlElement.SetAttribute("macro", "");
			base.TopNode.AppendChild(xmlElement);
			xmlElement.InnerXml = $"<xdr:nvGraphicFramePr><xdr:cNvPr id=\"{_id}\" name=\"Chart 1\" /><xdr:cNvGraphicFramePr /></xdr:nvGraphicFramePr><xdr:xfrm><a:off x=\"0\" y=\"0\" /> <a:ext cx=\"0\" cy=\"0\" /></xdr:xfrm><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rId1\" />   </a:graphicData>  </a:graphic>";
			base.TopNode.AppendChild(base.TopNode.OwnerDocument.CreateElement("clientData", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"));
			ZipPackage package = drawings.Worksheet._package.Package;
			UriChart = XmlHelper.GetNewUri(package, "/xl/charts/chart{0}.xml");
			ChartXml = new XmlDocument();
			ChartXml.PreserveWhitespace = false;
			XmlHelper.LoadXmlSafe(ChartXml, ChartStartXml(type), Encoding.UTF8);
			Part = package.CreatePart(UriChart, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml", _drawings._package.Compression);
			StreamWriter streamWriter = new StreamWriter(Part.GetStream(FileMode.Create, FileAccess.Write));
			ChartXml.Save(streamWriter);
			streamWriter.Close();
			package.Flush();
			ZipPackageRelationship zipPackageRelationship = drawings.Part.CreateRelationship(UriHelper.GetRelativeUri(drawings.UriDrawing, UriChart), TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart");
			xmlElement.SelectSingleNode("a:graphic/a:graphicData/c:chart", base.NameSpaceManager).Attributes["r:id"].Value = zipPackageRelationship.Id;
			package.Flush();
			_chartNode = ChartXml.SelectSingleNode($"c:chartSpace/c:chart/c:plotArea/{GetChartNodeText()}", base.NameSpaceManager);
		}
		else
		{
			ChartXml = topChart.ChartXml;
			Part = topChart.Part;
			_plotArea = topChart.PlotArea;
			UriChart = topChart.UriChart;
			_axis = topChart._axis;
			XmlNode chartNode = _plotArea.ChartTypes[_plotArea.ChartTypes.Count - 1].ChartNode;
			_chartNode = ChartXml.CreateElement(GetChartNodeText(), "http://schemas.openxmlformats.org/drawingml/2006/chart");
			chartNode.ParentNode.InsertAfter(_chartNode, chartNode);
			if (topChart.Axis.Length == 0)
			{
				AddAxis();
			}
			string chartSerieStartXml = GetChartSerieStartXml(type, int.Parse(topChart.Axis[0].Id), int.Parse(topChart.Axis[1].Id), (topChart.Axis.Length > 2) ? int.Parse(topChart.Axis[2].Id) : (-1));
			_chartNode.InnerXml = chartSerieStartXml;
		}
		GetPositionSize();
	}

	private void LoadAxis()
	{
		XmlNodeList xmlNodeList = _chartNode.SelectNodes("c:axId", base.NameSpaceManager);
		List<ExcelChartAxis> list = new List<ExcelChartAxis>();
		foreach (XmlNode item2 in xmlNodeList)
		{
			string value = item2.Attributes["val"].Value;
			XmlNodeList xmlNodeList2 = ChartXml.SelectNodes("c:chartSpace/c:chart/c:plotArea" + $"/*/c:axId[@val=\"{value}\"]", base.NameSpaceManager);
			if (xmlNodeList2 == null || xmlNodeList2.Count <= 1)
			{
				continue;
			}
			foreach (XmlNode item3 in xmlNodeList2)
			{
				if (item3.ParentNode.LocalName.EndsWith("Ax"))
				{
					XmlNode parentNode = xmlNodeList2[1].ParentNode;
					ExcelChartAxis item = new ExcelChartAxis(base.NameSpaceManager, parentNode);
					list.Add(item);
				}
			}
		}
		_axis = list.ToArray();
		if (_axis.Length != 0)
		{
			XAxis = _axis[0];
		}
		if (_axis.Length > 1)
		{
			YAxis = _axis[1];
		}
	}

	internal virtual eChartType GetChartType(string name)
	{
		switch (name)
		{
		case "area3DChart":
			if (Grouping == eGrouping.Stacked)
			{
				return eChartType.AreaStacked3D;
			}
			if (Grouping == eGrouping.PercentStacked)
			{
				return eChartType.AreaStacked1003D;
			}
			return eChartType.Area3D;
		case "areaChart":
			if (Grouping == eGrouping.Stacked)
			{
				return eChartType.AreaStacked;
			}
			if (Grouping == eGrouping.PercentStacked)
			{
				return eChartType.AreaStacked100;
			}
			return eChartType.Area;
		case "doughnutChart":
			return eChartType.Doughnut;
		case "pie3DChart":
			return eChartType.Pie3D;
		case "pieChart":
			return eChartType.Pie;
		case "radarChart":
			return eChartType.Radar;
		case "scatterChart":
			return eChartType.XYScatter;
		case "surface3DChart":
		case "surfaceChart":
			return eChartType.Surface;
		case "stockChart":
			return eChartType.StockHLC;
		default:
			return (eChartType)0;
		}
	}

	private string ChartStartXml(eChartType type)
	{
		StringBuilder stringBuilder = new StringBuilder();
		int num = 1;
		int num2 = 2;
		int num3 = (IsTypeSurface() ? 3 : (-1));
		stringBuilder.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
		stringBuilder.AppendFormat("<c:chartSpace xmlns:c=\"{0}\" xmlns:a=\"{1}\" xmlns:r=\"{2}\">", "http://schemas.openxmlformats.org/drawingml/2006/chart", "http://schemas.openxmlformats.org/drawingml/2006/main", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
		stringBuilder.Append("<c:chart>");
		stringBuilder.AppendFormat("{0}{1}<c:plotArea><c:layout/>", AddPerspectiveXml(type), AddSurfaceXml(type));
		string chartNodeText = GetChartNodeText();
		stringBuilder.AppendFormat("<{0}>", chartNodeText);
		stringBuilder.Append(GetChartSerieStartXml(type, num, num2, num3));
		stringBuilder.AppendFormat("</{0}>", chartNodeText);
		if (!IsTypePieDoughnut())
		{
			if (IsTypeScatterBubble())
			{
				stringBuilder.AppendFormat("<c:valAx><c:axId val=\"{0}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"{1}\"/><c:crosses val=\"autoZero\"/></c:valAx>", num, num2);
			}
			else
			{
				stringBuilder.AppendFormat("<c:catAx><c:axId val=\"{0}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"{1}\"/><c:crosses val=\"autoZero\"/><c:auto val=\"1\"/><c:lblAlgn val=\"ctr\"/><c:lblOffset val=\"100\"/></c:catAx>", num, num2);
			}
			stringBuilder.AppendFormat("<c:valAx><c:axId val=\"{1}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"l\"/><c:majorGridlines/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"{0}\"/><c:crosses val=\"autoZero\"/><c:crossBetween val=\"between\"/></c:valAx>", num, num2);
			if (num3 == 3)
			{
				stringBuilder.AppendFormat("<c:serAx><c:axId val=\"{0}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"{1}\"/><c:crosses val=\"autoZero\"/></c:serAx>", num3, num2);
			}
		}
		stringBuilder.AppendFormat("</c:plotArea><c:legend><c:legendPos val=\"r\"/><c:layout/><c:overlay val=\"0\" /></c:legend><c:plotVisOnly val=\"1\"/></c:chart>", num, num2);
		stringBuilder.Append("<c:printSettings><c:headerFooter/><c:pageMargins b=\"0.75\" l=\"0.7\" r=\"0.7\" t=\"0.75\" header=\"0.3\" footer=\"0.3\"/><c:pageSetup/></c:printSettings></c:chartSpace>");
		return stringBuilder.ToString();
	}

	private string GetChartSerieStartXml(eChartType type, int axID, int xAxID, int serAxID)
	{
		StringBuilder stringBuilder = new StringBuilder();
		stringBuilder.Append(AddScatterType(type));
		stringBuilder.Append(AddRadarType(type));
		stringBuilder.Append(AddBarDir(type));
		stringBuilder.Append(AddGrouping());
		stringBuilder.Append(AddVaryColors());
		stringBuilder.Append(AddHasMarker(type));
		stringBuilder.Append(AddShape(type));
		stringBuilder.Append(AddFirstSliceAng(type));
		stringBuilder.Append(AddHoleSize(type));
		if (ChartType == eChartType.BarStacked100 || ChartType == eChartType.BarStacked || ChartType == eChartType.ColumnStacked || ChartType == eChartType.ColumnStacked100)
		{
			stringBuilder.Append("<c:overlap val=\"100\"/>");
		}
		if (IsTypeSurface())
		{
			stringBuilder.Append("<c:bandFmts/>");
		}
		stringBuilder.Append(AddAxisId(axID, xAxID, serAxID));
		return stringBuilder.ToString();
	}

	private string AddAxisId(int axID, int xAxID, int serAxID)
	{
		if (!IsTypePieDoughnut())
		{
			if (IsTypeSurface())
			{
				return $"<c:axId val=\"{axID}\"/><c:axId val=\"{xAxID}\"/><c:axId val=\"{serAxID}\"/>";
			}
			return $"<c:axId val=\"{axID}\"/><c:axId val=\"{xAxID}\"/>";
		}
		return "";
	}

	private string AddAxType()
	{
		switch (ChartType)
		{
		case eChartType.XYScatter:
		case eChartType.Bubble:
		case eChartType.XYScatterSmooth:
		case eChartType.XYScatterSmoothNoMarkers:
		case eChartType.XYScatterLines:
		case eChartType.XYScatterLinesNoMarkers:
		case eChartType.Bubble3DEffect:
			return "valAx";
		default:
			return "catAx";
		}
	}

	private string AddScatterType(eChartType type)
	{
		if (type == eChartType.XYScatter || type == eChartType.XYScatterLines || type == eChartType.XYScatterLinesNoMarkers || type == eChartType.XYScatterSmooth || type == eChartType.XYScatterSmoothNoMarkers)
		{
			return "<c:scatterStyle val=\"\" />";
		}
		return "";
	}

	private string AddRadarType(eChartType type)
	{
		if (type == eChartType.Radar || type == eChartType.RadarFilled || type == eChartType.RadarMarkers)
		{
			return "<c:radarStyle val=\"\" />";
		}
		return "";
	}

	private string AddGrouping()
	{
		if (IsTypeShape() || IsTypeLine())
		{
			return "<c:grouping val=\"standard\"/>";
		}
		return "";
	}

	private string AddHoleSize(eChartType type)
	{
		if (type == eChartType.Doughnut || type == eChartType.DoughnutExploded)
		{
			return "<c:holeSize val=\"50\" />";
		}
		return "";
	}

	private string AddFirstSliceAng(eChartType type)
	{
		if (type == eChartType.Doughnut || type == eChartType.DoughnutExploded)
		{
			return "<c:firstSliceAng val=\"0\" />";
		}
		return "";
	}

	private string AddVaryColors()
	{
		if (IsTypePieDoughnut())
		{
			return "<c:varyColors val=\"1\" />";
		}
		return "<c:varyColors val=\"0\" />";
	}

	private string AddHasMarker(eChartType type)
	{
		if (type == eChartType.LineMarkers || type == eChartType.LineMarkersStacked || type == eChartType.LineMarkersStacked100)
		{
			return "<c:marker val=\"1\"/>";
		}
		return "";
	}

	private string AddShape(eChartType type)
	{
		if (IsTypeShape())
		{
			return "<c:shape val=\"box\" />";
		}
		return "";
	}

	private string AddBarDir(eChartType type)
	{
		if (IsTypeShape())
		{
			return "<c:barDir val=\"col\" />";
		}
		return "";
	}

	private string AddPerspectiveXml(eChartType type)
	{
		if (IsType3D())
		{
			return "<c:view3D><c:perspective val=\"30\" /></c:view3D>";
		}
		return "";
	}

	private string AddSurfaceXml(eChartType type)
	{
		if (IsTypeSurface())
		{
			return AddSurfacePart("floor") + AddSurfacePart("sideWall") + AddSurfacePart("backWall");
		}
		return "";
	}

	private string AddSurfacePart(string name)
	{
		return string.Format("<c:{0}><c:thickness val=\"0\"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/><a:sp3d/></c:spPr></c:{0}>", name);
	}

	internal static bool IsType3D(eChartType chartType)
	{
		if (chartType != eChartType.Area3D && chartType != eChartType.AreaStacked3D && chartType != eChartType.AreaStacked1003D && chartType != eChartType.BarClustered3D && chartType != eChartType.BarStacked3D && chartType != eChartType.BarStacked1003D && chartType != eChartType.Column3D && chartType != eChartType.ColumnClustered3D && chartType != eChartType.ColumnStacked3D && chartType != eChartType.ColumnStacked1003D && chartType != eChartType.Line3D && chartType != eChartType.Pie3D && chartType != eChartType.PieExploded3D && chartType != eChartType.ConeBarClustered && chartType != eChartType.ConeBarStacked && chartType != eChartType.ConeBarStacked100 && chartType != eChartType.ConeCol && chartType != eChartType.ConeColClustered && chartType != eChartType.ConeColStacked && chartType != eChartType.ConeColStacked100 && chartType != eChartType.CylinderBarClustered && chartType != eChartType.CylinderBarStacked && chartType != eChartType.CylinderBarStacked100 && chartType != eChartType.CylinderCol && chartType != eChartType.CylinderColClustered && chartType != eChartType.CylinderColStacked && chartType != eChartType.CylinderColStacked100 && chartType != eChartType.PyramidBarClustered && chartType != eChartType.PyramidBarStacked && chartType != eChartType.PyramidBarStacked100 && chartType != eChartType.PyramidCol && chartType != eChartType.PyramidColClustered && chartType != eChartType.PyramidColStacked && chartType != eChartType.PyramidColStacked100 && chartType != eChartType.Surface && chartType != eChartType.SurfaceTopView && chartType != eChartType.SurfaceTopViewWireframe)
		{
			return chartType == eChartType.SurfaceWireframe;
		}
		return true;
	}

	protected internal bool IsType3D()
	{
		return IsType3D(ChartType);
	}

	protected bool IsTypeLine()
	{
		if (ChartType != eChartType.Line && ChartType != eChartType.LineMarkers && ChartType != eChartType.LineMarkersStacked100 && ChartType != eChartType.LineStacked && ChartType != eChartType.LineStacked100)
		{
			return ChartType == eChartType.Line3D;
		}
		return true;
	}

	protected bool IsTypeScatterBubble()
	{
		if (ChartType != eChartType.XYScatter && ChartType != eChartType.XYScatterLines && ChartType != eChartType.XYScatterLinesNoMarkers && ChartType != eChartType.XYScatterSmooth && ChartType != eChartType.XYScatterSmoothNoMarkers && ChartType != eChartType.Bubble)
		{
			return ChartType == eChartType.Bubble3DEffect;
		}
		return true;
	}

	protected bool IsTypeSurface()
	{
		if (ChartType != eChartType.Surface && ChartType != eChartType.SurfaceTopView && ChartType != eChartType.SurfaceTopViewWireframe)
		{
			return ChartType == eChartType.SurfaceWireframe;
		}
		return true;
	}

	protected bool IsTypeShape()
	{
		if (ChartType != eChartType.BarClustered3D && ChartType != eChartType.BarStacked3D && ChartType != eChartType.BarStacked1003D && ChartType != eChartType.BarClustered3D && ChartType != eChartType.BarStacked3D && ChartType != eChartType.BarStacked1003D && ChartType != eChartType.Column3D && ChartType != eChartType.ColumnClustered3D && ChartType != eChartType.ColumnStacked3D && ChartType != eChartType.ColumnStacked1003D && ChartType != eChartType.ConeBarClustered && ChartType != eChartType.ConeBarStacked && ChartType != eChartType.ConeBarStacked100 && ChartType != eChartType.ConeCol && ChartType != eChartType.ConeColClustered && ChartType != eChartType.ConeColStacked && ChartType != eChartType.ConeColStacked100 && ChartType != eChartType.CylinderBarClustered && ChartType != eChartType.CylinderBarStacked && ChartType != eChartType.CylinderBarStacked100 && ChartType != eChartType.CylinderCol && ChartType != eChartType.CylinderColClustered && ChartType != eChartType.CylinderColStacked && ChartType != eChartType.CylinderColStacked100 && ChartType != eChartType.PyramidBarClustered && ChartType != eChartType.PyramidBarStacked && ChartType != eChartType.PyramidBarStacked100 && ChartType != eChartType.PyramidCol && ChartType != eChartType.PyramidColClustered && ChartType != eChartType.PyramidColStacked)
		{
			return ChartType == eChartType.PyramidColStacked100;
		}
		return true;
	}

	protected internal bool IsTypePercentStacked()
	{
		if (ChartType != eChartType.AreaStacked100 && ChartType != eChartType.BarStacked100 && ChartType != eChartType.BarStacked1003D && ChartType != eChartType.ColumnStacked100 && ChartType != eChartType.ColumnStacked1003D && ChartType != eChartType.ConeBarStacked100 && ChartType != eChartType.ConeColStacked100 && ChartType != eChartType.CylinderBarStacked100 && ChartType != eChartType.CylinderColStacked && ChartType != eChartType.LineMarkersStacked100 && ChartType != eChartType.LineStacked100 && ChartType != eChartType.PyramidBarStacked100)
		{
			return ChartType == eChartType.PyramidColStacked100;
		}
		return true;
	}

	protected internal bool IsTypeStacked()
	{
		if (ChartType != eChartType.AreaStacked && ChartType != eChartType.AreaStacked3D && ChartType != eChartType.BarStacked && ChartType != eChartType.BarStacked3D && ChartType != eChartType.ColumnStacked3D && ChartType != eChartType.ColumnStacked && ChartType != eChartType.ConeBarStacked && ChartType != eChartType.ConeColStacked && ChartType != eChartType.CylinderBarStacked && ChartType != eChartType.CylinderColStacked && ChartType != eChartType.LineMarkersStacked && ChartType != eChartType.LineStacked && ChartType != eChartType.PyramidBarStacked)
		{
			return ChartType == eChartType.PyramidColStacked;
		}
		return true;
	}

	protected bool IsTypeClustered()
	{
		if (ChartType != eChartType.BarClustered && ChartType != eChartType.BarClustered3D && ChartType != eChartType.ColumnClustered3D && ChartType != eChartType.ColumnClustered && ChartType != eChartType.ConeBarClustered && ChartType != eChartType.ConeColClustered && ChartType != eChartType.CylinderBarClustered && ChartType != eChartType.CylinderColClustered && ChartType != eChartType.PyramidBarClustered)
		{
			return ChartType == eChartType.PyramidColClustered;
		}
		return true;
	}

	protected internal bool IsTypePieDoughnut()
	{
		if (ChartType != eChartType.Pie && ChartType != eChartType.PieExploded && ChartType != eChartType.PieOfPie && ChartType != eChartType.Pie3D && ChartType != eChartType.PieExploded3D && ChartType != eChartType.BarOfPie && ChartType != eChartType.Doughnut)
		{
			return ChartType == eChartType.DoughnutExploded;
		}
		return true;
	}

	protected string GetChartNodeText()
	{
		switch (ChartType)
		{
		case eChartType.Area3D:
		case eChartType.AreaStacked3D:
		case eChartType.AreaStacked1003D:
			return "c:area3DChart";
		case eChartType.Area:
		case eChartType.AreaStacked:
		case eChartType.AreaStacked100:
			return "c:areaChart";
		case eChartType.ColumnClustered:
		case eChartType.ColumnStacked:
		case eChartType.ColumnStacked100:
		case eChartType.BarClustered:
		case eChartType.BarStacked:
		case eChartType.BarStacked100:
			return "c:barChart";
		case eChartType.ColumnClustered3D:
		case eChartType.ColumnStacked3D:
		case eChartType.ColumnStacked1003D:
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
			return "c:bar3DChart";
		case eChartType.Bubble:
		case eChartType.Bubble3DEffect:
			return "c:bubbleChart";
		case eChartType.Doughnut:
		case eChartType.DoughnutExploded:
			return "c:doughnutChart";
		case eChartType.Line:
		case eChartType.LineStacked:
		case eChartType.LineStacked100:
		case eChartType.LineMarkers:
		case eChartType.LineMarkersStacked:
		case eChartType.LineMarkersStacked100:
			return "c:lineChart";
		case eChartType.Line3D:
			return "c:line3DChart";
		case eChartType.Pie:
		case eChartType.PieExploded:
			return "c:pieChart";
		case eChartType.PieOfPie:
		case eChartType.BarOfPie:
			return "c:ofPieChart";
		case eChartType.Pie3D:
		case eChartType.PieExploded3D:
			return "c:pie3DChart";
		case eChartType.Radar:
		case eChartType.RadarMarkers:
		case eChartType.RadarFilled:
			return "c:radarChart";
		case eChartType.XYScatter:
		case eChartType.XYScatterSmooth:
		case eChartType.XYScatterSmoothNoMarkers:
		case eChartType.XYScatterLines:
		case eChartType.XYScatterLinesNoMarkers:
			return "c:scatterChart";
		case eChartType.Surface:
		case eChartType.SurfaceWireframe:
			return "c:surface3DChart";
		case eChartType.SurfaceTopView:
		case eChartType.SurfaceTopViewWireframe:
			return "c:surfaceChart";
		case eChartType.StockHLC:
			return "c:stockChart";
		default:
			throw new NotImplementedException("Chart type not implemented");
		}
	}

	internal void AddAxis()
	{
		XmlElement xmlElement = ChartXml.CreateElement($"c:{AddAxType()}", "http://schemas.openxmlformats.org/drawingml/2006/chart");
		int num;
		if (_axis.Length == 0)
		{
			_plotArea.TopNode.AppendChild(xmlElement);
			num = 1;
		}
		else
		{
			_axis[0].TopNode.ParentNode.InsertAfter(xmlElement, _axis[_axis.Length - 1].TopNode);
			num = ((int.Parse(_axis[0].Id) < int.Parse(_axis[1].Id)) ? (int.Parse(_axis[1].Id) + 1) : (int.Parse(_axis[0].Id) + 1));
		}
		XmlElement xmlElement2 = ChartXml.CreateElement("c:valAx", "http://schemas.openxmlformats.org/drawingml/2006/chart");
		xmlElement.ParentNode.InsertAfter(xmlElement2, xmlElement);
		if (_axis.Length == 0)
		{
			xmlElement.InnerXml = $"<c:axId val=\"{num}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\" /><c:axPos val=\"b\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"{num + 1}\"/><c:crosses val=\"autoZero\"/><c:auto val=\"1\"/><c:lblAlgn val=\"ctr\"/><c:lblOffset val=\"100\"/>";
			xmlElement2.InnerXml = string.Format("<c:axId val=\"{1}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\" /><c:axPos val=\"l\"/><c:majorGridlines/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"{0}\"/><c:crosses val=\"autoZero\"/><c:crossBetween val=\"between\"/>", num, num + 1);
		}
		else
		{
			xmlElement.InnerXml = $"<c:axId val=\"{num}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"1\" /><c:axPos val=\"b\"/><c:tickLblPos val=\"none\"/><c:crossAx val=\"{num + 1}\"/><c:crosses val=\"autoZero\"/>";
			xmlElement2.InnerXml = $"<c:axId val=\"{num + 1}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\" /><c:axPos val=\"r\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"{num}\"/><c:crosses val=\"max\"/><c:crossBetween val=\"between\"/>";
		}
		if (_axis.Length == 0)
		{
			_axis = new ExcelChartAxis[2];
		}
		else
		{
			ExcelChartAxis[] array = new ExcelChartAxis[_axis.Length + 2];
			Array.Copy(_axis, array, _axis.Length);
			_axis = array;
		}
		_axis[_axis.Length - 2] = new ExcelChartAxis(base.NameSpaceManager, xmlElement);
		_axis[_axis.Length - 1] = new ExcelChartAxis(base.NameSpaceManager, xmlElement2);
		foreach (ExcelChart item in (IEnumerable<ExcelChart>)_plotArea.ChartTypes)
		{
			item._axis = _axis;
		}
	}

	internal void RemoveSecondaryAxis()
	{
		throw new NotImplementedException("Not yet implemented");
	}

	private bool HasPrimaryAxis()
	{
		if (_plotArea.ChartTypes.Count == 1)
		{
			return false;
		}
		foreach (ExcelChart item in (IEnumerable<ExcelChart>)_plotArea.ChartTypes)
		{
			if (item != this && !item.UseSecondaryAxis && !item.IsTypePieDoughnut())
			{
				return true;
			}
		}
		return false;
	}

	private void CheckRemoveAxis(ExcelChartAxis excelChartAxis)
	{
		if (!ExistsAxis(excelChartAxis))
		{
			return;
		}
		ExcelChartAxis[] array = new ExcelChartAxis[Axis.Length - 1];
		int num = 0;
		ExcelChartAxis[] axis = Axis;
		foreach (ExcelChartAxis excelChartAxis2 in axis)
		{
			if (excelChartAxis2 != excelChartAxis)
			{
				array[num] = excelChartAxis2;
			}
		}
		foreach (ExcelChart item in (IEnumerable<ExcelChart>)_plotArea.ChartTypes)
		{
			item._axis = array;
		}
	}

	private bool ExistsAxis(ExcelChartAxis excelChartAxis)
	{
		foreach (ExcelChart item in (IEnumerable<ExcelChart>)_plotArea.ChartTypes)
		{
			if (item != this && (item.XAxis.AxisPosition == excelChartAxis.AxisPosition || item.YAxis.AxisPosition == excelChartAxis.AxisPosition))
			{
				return true;
			}
		}
		return false;
	}

	private string GetGroupingText(eGrouping grouping)
	{
		return grouping switch
		{
			eGrouping.Clustered => "clustered", 
			eGrouping.Stacked => "stacked", 
			eGrouping.PercentStacked => "percentStacked", 
			_ => "standard", 
		};
	}

	private eGrouping GetGroupingEnum(string grouping)
	{
		if (!(grouping == "stacked"))
		{
			if (grouping == "percentStacked")
			{
				return eGrouping.PercentStacked;
			}
			return eGrouping.Clustered;
		}
		return eGrouping.Stacked;
	}

	internal static ExcelChart GetChart(ExcelDrawings drawings, XmlNode node)
	{
		XmlNode xmlNode = node.SelectSingleNode("xdr:graphicFrame/a:graphic/a:graphicData/c:chart", drawings.NameSpaceManager);
		if (xmlNode != null)
		{
			ZipPackageRelationship relationship = drawings.Part.GetRelationship(xmlNode.Attributes["r:id"].Value);
			Uri uri = UriHelper.ResolvePartUri(drawings.UriDrawing, relationship.TargetUri);
			ZipPackagePart part = drawings.Part.Package.GetPart(uri);
			XmlDocument xmlDocument = new XmlDocument();
			XmlHelper.LoadXmlSafe(xmlDocument, part.GetStream());
			ExcelChart excelChart = null;
			{
				foreach (XmlElement childNode in xmlDocument.SelectSingleNode("c:chartSpace/c:chart/c:plotArea", drawings.NameSpaceManager).ChildNodes)
				{
					if (excelChart == null)
					{
						excelChart = GetChart(childNode, drawings, node, uri, part, xmlDocument, null);
						excelChart?.PlotArea.ChartTypes.Add(excelChart);
						continue;
					}
					ExcelChart chart = GetChart(childNode, null, null, null, null, null, excelChart);
					if (chart != null)
					{
						excelChart.PlotArea.ChartTypes.Add(chart);
					}
				}
				return excelChart;
			}
		}
		return null;
	}

	internal static ExcelChart GetChart(XmlElement chartNode, ExcelDrawings drawings, XmlNode node, Uri uriChart, ZipPackagePart part, XmlDocument chartXml, ExcelChart topChart)
	{
		switch (chartNode.LocalName)
		{
		case "area3DChart":
		case "areaChart":
		case "stockChart":
			if (topChart == null)
			{
				return new ExcelChart(drawings, node, uriChart, part, chartXml, chartNode);
			}
			return new ExcelChart(topChart, chartNode);
		case "surface3DChart":
		case "surfaceChart":
			if (topChart == null)
			{
				return new ExcelSurfaceChart(drawings, node, uriChart, part, chartXml, chartNode);
			}
			return new ExcelSurfaceChart(topChart, chartNode);
		case "radarChart":
			if (topChart == null)
			{
				return new ExcelRadarChart(drawings, node, uriChart, part, chartXml, chartNode);
			}
			return new ExcelRadarChart(topChart, chartNode);
		case "bubbleChart":
			if (topChart == null)
			{
				return new ExcelBubbleChart(drawings, node, uriChart, part, chartXml, chartNode);
			}
			return new ExcelBubbleChart(topChart, chartNode);
		case "barChart":
		case "bar3DChart":
			if (topChart == null)
			{
				return new ExcelBarChart(drawings, node, uriChart, part, chartXml, chartNode);
			}
			return new ExcelBarChart(topChart, chartNode);
		case "doughnutChart":
			if (topChart == null)
			{
				return new ExcelDoughnutChart(drawings, node, uriChart, part, chartXml, chartNode);
			}
			return new ExcelDoughnutChart(topChart, chartNode);
		case "pie3DChart":
		case "pieChart":
			if (topChart == null)
			{
				return new ExcelPieChart(drawings, node, uriChart, part, chartXml, chartNode);
			}
			return new ExcelPieChart(topChart, chartNode);
		case "ofPieChart":
			if (topChart == null)
			{
				return new ExcelOfPieChart(drawings, node, uriChart, part, chartXml, chartNode);
			}
			return new ExcelBarChart(topChart, chartNode);
		case "lineChart":
		case "line3DChart":
			if (topChart == null)
			{
				return new ExcelLineChart(drawings, node, uriChart, part, chartXml, chartNode);
			}
			return new ExcelLineChart(topChart, chartNode);
		case "scatterChart":
			if (topChart == null)
			{
				return new ExcelScatterChart(drawings, node, uriChart, part, chartXml, chartNode);
			}
			return new ExcelScatterChart(topChart, chartNode);
		default:
			return null;
		}
	}

	internal static ExcelChart GetNewChart(ExcelDrawings drawings, XmlNode drawNode, eChartType chartType, ExcelChart topChart, ExcelPivotTable PivotTableSource)
	{
		switch (chartType)
		{
		case eChartType.Pie3D:
		case eChartType.Pie:
		case eChartType.PieExploded:
		case eChartType.PieExploded3D:
			return new ExcelPieChart(drawings, drawNode, chartType, topChart, PivotTableSource);
		case eChartType.PieOfPie:
		case eChartType.BarOfPie:
			return new ExcelOfPieChart(drawings, drawNode, chartType, topChart, PivotTableSource);
		case eChartType.Doughnut:
		case eChartType.DoughnutExploded:
			return new ExcelDoughnutChart(drawings, drawNode, chartType, topChart, PivotTableSource);
		case eChartType.Column3D:
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
			return new ExcelBarChart(drawings, drawNode, chartType, topChart, PivotTableSource);
		case eChartType.XYScatter:
		case eChartType.XYScatterSmooth:
		case eChartType.XYScatterSmoothNoMarkers:
		case eChartType.XYScatterLines:
		case eChartType.XYScatterLinesNoMarkers:
			return new ExcelScatterChart(drawings, drawNode, chartType, topChart, PivotTableSource);
		case eChartType.Line3D:
		case eChartType.Line:
		case eChartType.LineStacked:
		case eChartType.LineStacked100:
		case eChartType.LineMarkers:
		case eChartType.LineMarkersStacked:
		case eChartType.LineMarkersStacked100:
			return new ExcelLineChart(drawings, drawNode, chartType, topChart, PivotTableSource);
		case eChartType.Bubble:
		case eChartType.Bubble3DEffect:
			return new ExcelBubbleChart(drawings, drawNode, chartType, topChart, PivotTableSource);
		case eChartType.Radar:
		case eChartType.RadarMarkers:
		case eChartType.RadarFilled:
			return new ExcelRadarChart(drawings, drawNode, chartType, topChart, PivotTableSource);
		case eChartType.Surface:
		case eChartType.SurfaceWireframe:
		case eChartType.SurfaceTopView:
		case eChartType.SurfaceTopViewWireframe:
			return new ExcelSurfaceChart(drawings, drawNode, chartType, topChart, PivotTableSource);
		default:
			return new ExcelChart(drawings, drawNode, chartType, topChart, PivotTableSource);
		}
	}

	internal void SetPivotSource(ExcelPivotTable pivotTableSource)
	{
		PivotTableSource = pivotTableSource;
		XmlElement xmlElement = ChartXml.SelectSingleNode("c:chartSpace/c:chart", base.NameSpaceManager) as XmlElement;
		XmlElement xmlElement2 = ChartXml.CreateElement("pivotSource", "http://schemas.openxmlformats.org/drawingml/2006/chart");
		xmlElement.ParentNode.InsertBefore(xmlElement2, xmlElement);
		xmlElement2.InnerXml = $"<c:name>[]{PivotTableSource.WorkSheet.Name}!{pivotTableSource.Name}</c:name><c:fmtId val=\"0\"/>";
		XmlElement xmlElement3 = ChartXml.CreateElement("pivotFmts", "http://schemas.openxmlformats.org/drawingml/2006/chart");
		xmlElement.PrependChild(xmlElement3);
		xmlElement3.InnerXml = "<c:pivotFmt><c:idx val=\"0\"/><c:marker><c:symbol val=\"none\"/></c:marker></c:pivotFmt>";
		Series.AddPivotSerie(pivotTableSource);
	}

	internal override void DeleteMe()
	{
		try
		{
			Part.Package.DeletePart(UriChart);
		}
		catch (Exception innerException)
		{
			throw new InvalidDataException("EPPlus internal error when deleteing chart.", innerException);
		}
		base.DeleteMe();
	}
}
