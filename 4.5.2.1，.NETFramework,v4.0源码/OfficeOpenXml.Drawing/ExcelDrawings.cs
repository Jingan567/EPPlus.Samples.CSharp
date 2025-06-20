using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Xml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Drawing;

public class ExcelDrawings : IEnumerable<ExcelDrawing>, IEnumerable, IDisposable
{
	internal class ImageCompare
	{
		internal byte[] image { get; set; }

		internal string relID { get; set; }

		internal bool Comparer(byte[] compareImg)
		{
			if (compareImg.Length != image.Length)
			{
				return false;
			}
			for (int i = 0; i < image.Length; i++)
			{
				if (image[i] != compareImg[i])
				{
					return false;
				}
			}
			return true;
		}
	}

	private XmlDocument _drawingsXml = new XmlDocument();

	private Dictionary<string, int> _drawingNames;

	private List<ExcelDrawing> _drawings;

	internal Dictionary<string, string> _hashes = new Dictionary<string, string>();

	internal ExcelPackage _package;

	internal ZipPackageRelationship _drawingRelation;

	private XmlNamespaceManager _nsManager;

	private ZipPackagePart _part;

	private Uri _uriDrawing;

	internal ExcelWorksheet Worksheet { get; set; }

	public XmlDocument DrawingXml => _drawingsXml;

	public XmlNamespaceManager NameSpaceManager => _nsManager;

	public ExcelDrawing this[int PositionID] => _drawings[PositionID];

	public ExcelDrawing this[string Name]
	{
		get
		{
			if (_drawingNames.ContainsKey(Name))
			{
				return _drawings[_drawingNames[Name]];
			}
			return null;
		}
	}

	public int Count
	{
		get
		{
			if (_drawings == null)
			{
				return 0;
			}
			return _drawings.Count;
		}
	}

	internal ZipPackagePart Part => _part;

	public Uri UriDrawing => _uriDrawing;

	internal ExcelDrawings(ExcelPackage xlPackage, ExcelWorksheet sheet)
	{
		_drawingsXml = new XmlDocument();
		_drawingsXml.PreserveWhitespace = false;
		_drawings = new List<ExcelDrawing>();
		_drawingNames = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
		_package = xlPackage;
		Worksheet = sheet;
		XmlNode xmlNode = sheet.WorksheetXml.SelectSingleNode("//d:drawing", sheet.NameSpaceManager);
		CreateNSM();
		if (xmlNode != null)
		{
			_drawingRelation = sheet.Part.GetRelationship(xmlNode.Attributes["r:id"].Value);
			_uriDrawing = UriHelper.ResolvePartUri(sheet.WorksheetUri, _drawingRelation.TargetUri);
			_part = xlPackage.Package.GetPart(_uriDrawing);
			XmlHelper.LoadXmlSafe(_drawingsXml, _part.GetStream());
			AddDrawings();
		}
	}

	private void AddDrawings()
	{
		foreach (XmlNode item in _drawingsXml.SelectNodes("//*[self::xdr:twoCellAnchor or self::xdr:oneCellAnchor or self::xdr:absoluteAnchor]", NameSpaceManager))
		{
			ExcelDrawing excelDrawing = item.LocalName switch
			{
				"oneCellAnchor" => ExcelDrawing.GetDrawing(this, item), 
				"twoCellAnchor" => ExcelDrawing.GetDrawing(this, item), 
				"absoluteAnchor" => ExcelDrawing.GetDrawing(this, item), 
				_ => null, 
			};
			if (excelDrawing != null)
			{
				_drawings.Add(excelDrawing);
				if (!_drawingNames.ContainsKey(excelDrawing.Name))
				{
					_drawingNames.Add(excelDrawing.Name, _drawings.Count - 1);
				}
			}
		}
	}

	private void CreateNSM()
	{
		NameTable nameTable = new NameTable();
		_nsManager = new XmlNamespaceManager(nameTable);
		_nsManager.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
		_nsManager.AddNamespace("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
		_nsManager.AddNamespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
		_nsManager.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
	}

	public IEnumerator GetEnumerator()
	{
		return _drawings.GetEnumerator();
	}

	IEnumerator<ExcelDrawing> IEnumerable<ExcelDrawing>.GetEnumerator()
	{
		return _drawings.GetEnumerator();
	}

	public ExcelChart AddChart(string Name, eChartType ChartType, ExcelPivotTable PivotTableSource)
	{
		if (_drawingNames.ContainsKey(Name))
		{
			throw new Exception("Name already exists in the drawings collection");
		}
		if (ChartType == eChartType.StockHLC || ChartType == eChartType.StockOHLC || ChartType == eChartType.StockVOHLC)
		{
			throw new NotImplementedException("Chart type is not supported in the current version");
		}
		if (Worksheet is ExcelChartsheet && _drawings.Count > 0)
		{
			throw new InvalidOperationException("Chart Worksheets can't have more than one chart");
		}
		XmlElement drawNode = CreateDrawingXml();
		ExcelChart newChart = ExcelChart.GetNewChart(this, drawNode, ChartType, null, PivotTableSource);
		newChart.Name = Name;
		_drawings.Add(newChart);
		_drawingNames.Add(Name, _drawings.Count - 1);
		return newChart;
	}

	public ExcelChart AddChart(string Name, eChartType ChartType)
	{
		return AddChart(Name, ChartType, null);
	}

	public ExcelPicture AddPicture(string Name, Image image)
	{
		return AddPicture(Name, image, null);
	}

	public ExcelPicture AddPicture(string Name, Image image, Uri Hyperlink)
	{
		if (image != null)
		{
			if (_drawingNames.ContainsKey(Name))
			{
				throw new Exception("Name already exists in the drawings collection");
			}
			XmlElement xmlElement = CreateDrawingXml();
			xmlElement.SetAttribute("editAs", "oneCell");
			ExcelPicture excelPicture = new ExcelPicture(this, xmlElement, image, Hyperlink);
			excelPicture.Name = Name;
			_drawings.Add(excelPicture);
			_drawingNames.Add(Name, _drawings.Count - 1);
			return excelPicture;
		}
		throw new Exception("AddPicture: Image can't be null");
	}

	public ExcelPicture AddPicture(string Name, FileInfo ImageFile)
	{
		return AddPicture(Name, ImageFile, null);
	}

	public ExcelPicture AddPicture(string Name, FileInfo ImageFile, Uri Hyperlink)
	{
		if (Worksheet is ExcelChartsheet && _drawings.Count > 0)
		{
			throw new InvalidOperationException("Chart worksheets can't have more than one drawing");
		}
		if (ImageFile != null)
		{
			if (_drawingNames.ContainsKey(Name))
			{
				throw new Exception("Name already exists in the drawings collection");
			}
			XmlElement xmlElement = CreateDrawingXml();
			xmlElement.SetAttribute("editAs", "oneCell");
			ExcelPicture excelPicture = new ExcelPicture(this, xmlElement, ImageFile, Hyperlink);
			excelPicture.Name = Name;
			_drawings.Add(excelPicture);
			_drawingNames.Add(Name, _drawings.Count - 1);
			return excelPicture;
		}
		throw new Exception("AddPicture: ImageFile can't be null");
	}

	public ExcelShape AddShape(string Name, eShapeStyle Style)
	{
		if (Worksheet is ExcelChartsheet && _drawings.Count > 0)
		{
			throw new InvalidOperationException("Chart worksheets can't have more than one drawing");
		}
		if (_drawingNames.ContainsKey(Name))
		{
			throw new Exception("Name already exists in the drawings collection");
		}
		XmlElement node = CreateDrawingXml();
		ExcelShape excelShape = new ExcelShape(this, node, Style);
		excelShape.Name = Name;
		excelShape.Style = Style;
		_drawings.Add(excelShape);
		_drawingNames.Add(Name, _drawings.Count - 1);
		return excelShape;
	}

	public ExcelShape AddShape(string Name, ExcelShape Source)
	{
		if (Worksheet is ExcelChartsheet && _drawings.Count > 0)
		{
			throw new InvalidOperationException("Chart worksheets can't have more than one drawing");
		}
		if (_drawingNames.ContainsKey(Name))
		{
			throw new Exception("Name already exists in the drawings collection");
		}
		XmlElement xmlElement = CreateDrawingXml();
		xmlElement.InnerXml = Source.TopNode.InnerXml;
		ExcelShape excelShape = new ExcelShape(this, xmlElement);
		excelShape.Name = Name;
		excelShape.Style = Source.Style;
		_drawings.Add(excelShape);
		_drawingNames.Add(Name, _drawings.Count - 1);
		return excelShape;
	}

	private XmlElement CreateDrawingXml()
	{
		if (DrawingXml.DocumentElement == null)
		{
			DrawingXml.LoadXml(string.Format("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><xdr:wsDr xmlns:xdr=\"{0}\" xmlns:a=\"{1}\" />", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing", "http://schemas.openxmlformats.org/drawingml/2006/main"));
			ZipPackage package = Worksheet._package.Package;
			int sheetID = Worksheet.SheetID;
			do
			{
				_uriDrawing = new Uri($"/xl/drawings/drawing{sheetID++}.xml", UriKind.Relative);
			}
			while (package.PartExists(_uriDrawing));
			_part = package.CreatePart(_uriDrawing, "application/vnd.openxmlformats-officedocument.drawing+xml", _package.Compression);
			StreamWriter streamWriter = new StreamWriter(_part.GetStream(FileMode.Create, FileAccess.Write));
			DrawingXml.Save(streamWriter);
			streamWriter.Close();
			package.Flush();
			_drawingRelation = Worksheet.Part.CreateRelationship(UriHelper.GetRelativeUri(Worksheet.WorksheetUri, _uriDrawing), TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing");
			((XmlElement)Worksheet.CreateNode("d:drawing")).SetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", _drawingRelation.Id);
			package.Flush();
		}
		XmlNode xmlNode = _drawingsXml.SelectSingleNode("//xdr:wsDr", NameSpaceManager);
		XmlElement xmlElement;
		if (Worksheet is ExcelChartsheet)
		{
			xmlElement = _drawingsXml.CreateElement("xdr", "absoluteAnchor", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
			XmlElement xmlElement2 = _drawingsXml.CreateElement("xdr", "pos", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
			xmlElement2.SetAttribute("y", "0");
			xmlElement2.SetAttribute("x", "0");
			xmlElement.AppendChild(xmlElement2);
			XmlElement xmlElement3 = _drawingsXml.CreateElement("xdr", "ext", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
			xmlElement3.SetAttribute("cy", "6072876");
			xmlElement3.SetAttribute("cx", "9299263");
			xmlElement.AppendChild(xmlElement3);
			xmlNode.AppendChild(xmlElement);
		}
		else
		{
			xmlElement = _drawingsXml.CreateElement("xdr", "twoCellAnchor", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
			xmlNode.AppendChild(xmlElement);
			XmlElement xmlElement4 = _drawingsXml.CreateElement("xdr", "from", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
			xmlElement.AppendChild(xmlElement4);
			xmlElement4.InnerXml = "<xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff>";
			XmlElement xmlElement5 = _drawingsXml.CreateElement("xdr", "to", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
			xmlElement.AppendChild(xmlElement5);
			xmlElement5.InnerXml = "<xdr:col>10</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff>";
		}
		return xmlElement;
	}

	public void Remove(int Index)
	{
		if (Worksheet is ExcelChartsheet && _drawings.Count > 0)
		{
			throw new InvalidOperationException("Can' remove charts from chart worksheets");
		}
		RemoveDrawing(Index);
	}

	internal void RemoveDrawing(int Index)
	{
		ExcelDrawing excelDrawing = _drawings[Index];
		excelDrawing.DeleteMe();
		for (int i = Index + 1; i < _drawings.Count; i++)
		{
			_drawingNames[_drawings[i].Name]--;
		}
		_drawingNames.Remove(excelDrawing.Name);
		_drawings.Remove(excelDrawing);
	}

	public void Remove(ExcelDrawing Drawing)
	{
		Remove(_drawingNames[Drawing.Name]);
	}

	public void Remove(string Name)
	{
		Remove(_drawingNames[Name]);
	}

	public void Clear()
	{
		if (Worksheet is ExcelChartsheet && _drawings.Count > 0)
		{
			throw new InvalidOperationException("Can' remove charts from chart worksheets");
		}
		ClearDrawings();
	}

	internal void ClearDrawings()
	{
		while (Count > 0)
		{
			RemoveDrawing(0);
		}
	}

	internal void AdjustWidth(int[,] pos)
	{
		int num = 0;
		IEnumerator enumerator = GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				ExcelDrawing excelDrawing = (ExcelDrawing)enumerator.Current;
				if (excelDrawing.EditAs != eEditAs.TwoCell)
				{
					if (excelDrawing.EditAs == eEditAs.Absolute)
					{
						excelDrawing.SetPixelLeft(pos[num, 0]);
					}
					excelDrawing.SetPixelWidth(pos[num, 1]);
				}
				num++;
			}
		}
		finally
		{
			IDisposable disposable = enumerator as IDisposable;
			if (disposable != null)
			{
				disposable.Dispose();
			}
		}
	}

	internal void AdjustHeight(int[,] pos)
	{
		int num = 0;
		IEnumerator enumerator = GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				ExcelDrawing excelDrawing = (ExcelDrawing)enumerator.Current;
				if (excelDrawing.EditAs != eEditAs.TwoCell)
				{
					if (excelDrawing.EditAs == eEditAs.Absolute)
					{
						excelDrawing.SetPixelTop(pos[num, 0]);
					}
					excelDrawing.SetPixelHeight(pos[num, 1]);
				}
				num++;
			}
		}
		finally
		{
			IDisposable disposable = enumerator as IDisposable;
			if (disposable != null)
			{
				disposable.Dispose();
			}
		}
	}

	internal int[,] GetDrawingWidths()
	{
		int[,] array = new int[Count, 2];
		int num = 0;
		IEnumerator enumerator = GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				ExcelDrawing excelDrawing = (ExcelDrawing)enumerator.Current;
				array[num, 0] = excelDrawing.GetPixelLeft();
				array[num++, 1] = excelDrawing.GetPixelWidth();
			}
			return array;
		}
		finally
		{
			IDisposable disposable = enumerator as IDisposable;
			if (disposable != null)
			{
				disposable.Dispose();
			}
		}
	}

	internal int[,] GetDrawingHeight()
	{
		int[,] array = new int[Count, 2];
		int num = 0;
		IEnumerator enumerator = GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				ExcelDrawing excelDrawing = (ExcelDrawing)enumerator.Current;
				array[num, 0] = excelDrawing.GetPixelTop();
				array[num++, 1] = excelDrawing.GetPixelHeight();
			}
			return array;
		}
		finally
		{
			IDisposable disposable = enumerator as IDisposable;
			if (disposable != null)
			{
				disposable.Dispose();
			}
		}
	}

	public void Dispose()
	{
		_drawingsXml = null;
		_hashes.Clear();
		_hashes = null;
		_part = null;
		_drawingNames.Clear();
		_drawingNames = null;
		_drawingRelation = null;
		foreach (ExcelDrawing drawing in _drawings)
		{
			drawing.Dispose();
		}
		_drawings.Clear();
		_drawings = null;
	}
}
