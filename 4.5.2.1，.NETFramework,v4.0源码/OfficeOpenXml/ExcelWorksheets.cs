using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils;
using OfficeOpenXml.VBA;

namespace OfficeOpenXml;

public class ExcelWorksheets : XmlHelper, IEnumerable<ExcelWorksheet>, IEnumerable, IDisposable
{
	private ExcelPackage _pck;

	private Dictionary<int, ExcelWorksheet> _worksheets;

	private XmlNamespaceManager _namespaceManager;

	private const string ERR_DUP_WORKSHEET = "A worksheet with this name already exists in the workbook";

	internal const string WORKSHEET_CONTENTTYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";

	internal const string CHARTSHEET_CONTENTTYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml";

	public int Count => _worksheets.Count;

	public ExcelWorksheet this[int PositionID]
	{
		get
		{
			if (_worksheets.ContainsKey(PositionID))
			{
				return _worksheets[PositionID];
			}
			throw new IndexOutOfRangeException("Worksheet position out of range.");
		}
	}

	public ExcelWorksheet this[string Name] => GetByName(Name);

	internal ExcelWorksheets(ExcelPackage pck, XmlNamespaceManager nsm, XmlNode topNode)
		: base(nsm, topNode)
	{
		_pck = pck;
		_namespaceManager = nsm;
		_worksheets = new Dictionary<int, ExcelWorksheet>();
		int num = _pck._worksheetAdd;
		foreach (XmlNode childNode in topNode.ChildNodes)
		{
			if (childNode.NodeType == XmlNodeType.Element)
			{
				string value = childNode.Attributes["name"].Value;
				string value2 = childNode.Attributes.GetNamedItem("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships").Value;
				int sheetID = Convert.ToInt32(childNode.Attributes["sheetId"].Value);
				eWorkSheetHidden eWorkSheetHidden2 = eWorkSheetHidden.Visible;
				XmlNode xmlNode2 = childNode.Attributes["state"];
				if (xmlNode2 != null)
				{
					eWorkSheetHidden2 = TranslateHidden(xmlNode2.Value);
				}
				ZipPackageRelationship relationship = pck.Workbook.Part.GetRelationship(value2);
				Uri uriWorksheet = UriHelper.ResolvePartUri(pck.Workbook.WorkbookUri, relationship.TargetUri);
				if (relationship.RelationshipType.EndsWith("chartsheet"))
				{
					_worksheets.Add(num, new ExcelChartsheet(_namespaceManager, _pck, value2, uriWorksheet, value, sheetID, num, eWorkSheetHidden2));
				}
				else
				{
					_worksheets.Add(num, new ExcelWorksheet(_namespaceManager, _pck, value2, uriWorksheet, value, sheetID, num, eWorkSheetHidden2));
				}
				num++;
			}
		}
	}

	private eWorkSheetHidden TranslateHidden(string value)
	{
		if (!(value == "hidden"))
		{
			if (value == "veryHidden")
			{
				return eWorkSheetHidden.VeryHidden;
			}
			return eWorkSheetHidden.Visible;
		}
		return eWorkSheetHidden.Hidden;
	}

	public IEnumerator<ExcelWorksheet> GetEnumerator()
	{
		return _worksheets.Values.GetEnumerator();
	}

	IEnumerator IEnumerable.GetEnumerator()
	{
		return _worksheets.Values.GetEnumerator();
	}

	public ExcelWorksheet Add(string Name)
	{
		return AddSheet(Name, isChart: false, null);
	}

	private ExcelWorksheet AddSheet(string Name, bool isChart, eChartType? chartType)
	{
		lock (_worksheets)
		{
			Name = ValidateFixSheetName(Name);
			if (GetByName(Name) != null)
			{
				throw new InvalidOperationException("A worksheet with this name already exists in the workbook : " + Name);
			}
			GetSheetURI(ref Name, out var sheetID, out var uriWorksheet, isChart);
			StreamWriter writer = new StreamWriter(_pck.Package.CreatePart(uriWorksheet, isChart ? "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml" : "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", _pck.Compression).GetStream(FileMode.Create, FileAccess.Write));
			CreateNewWorksheet(isChart).Save(writer);
			_pck.Package.Flush();
			string relID = CreateWorkbookRel(Name, sheetID, uriWorksheet, isChart);
			int num = _worksheets.Count + _pck._worksheetAdd;
			ExcelWorksheet excelWorksheet = ((!isChart) ? new ExcelWorksheet(_namespaceManager, _pck, relID, uriWorksheet, Name, sheetID, num, eWorkSheetHidden.Visible) : new ExcelChartsheet(_namespaceManager, _pck, relID, uriWorksheet, Name, sheetID, num, eWorkSheetHidden.Visible, chartType.Value));
			_worksheets.Add(num, excelWorksheet);
			if (_pck.Workbook.VbaProject != null)
			{
				string moduleNameFromWorksheet = _pck.Workbook.VbaProject.GetModuleNameFromWorksheet(excelWorksheet);
				_pck.Workbook.VbaProject.Modules.Add(new ExcelVBAModule(excelWorksheet.CodeNameChange)
				{
					Name = moduleNameFromWorksheet,
					Code = "",
					Attributes = _pck.Workbook.VbaProject.GetDocumentAttributes(Name, "0{00020820-0000-0000-C000-000000000046}"),
					Type = eModuleType.Document,
					HelpContext = 0
				});
				excelWorksheet.CodeModuleName = moduleNameFromWorksheet;
			}
			return excelWorksheet;
		}
	}

	public ExcelWorksheet Add(string Name, ExcelWorksheet Copy)
	{
		lock (_worksheets)
		{
			if (Copy is ExcelChartsheet)
			{
				throw new ArgumentException("Can not copy a chartsheet");
			}
			if (GetByName(Name) != null)
			{
				throw new InvalidOperationException("A worksheet with this name already exists in the workbook");
			}
			GetSheetURI(ref Name, out var sheetID, out var uriWorksheet, isChart: false);
			StreamWriter writer = new StreamWriter(_pck.Package.CreatePart(uriWorksheet, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", _pck.Compression).GetStream(FileMode.Create, FileAccess.Write));
			XmlDocument xmlDocument = new XmlDocument();
			xmlDocument.LoadXml(Copy.WorksheetXml.OuterXml);
			xmlDocument.Save(writer);
			_pck.Package.Flush();
			string relID = CreateWorkbookRel(Name, sheetID, uriWorksheet, isChart: false);
			ExcelWorksheet excelWorksheet = new ExcelWorksheet(_namespaceManager, _pck, relID, uriWorksheet, Name, sheetID, _worksheets.Count + _pck._worksheetAdd, eWorkSheetHidden.Visible);
			if (Copy.Comments.Count > 0)
			{
				CopyComment(Copy, excelWorksheet);
			}
			else if (Copy.VmlDrawingsComments.Count > 0)
			{
				CopyVmlDrawing(Copy, excelWorksheet);
			}
			CopyHeaderFooterPictures(Copy, excelWorksheet);
			if (Copy.Drawings.Count > 0)
			{
				CopyDrawing(Copy, excelWorksheet);
			}
			if (Copy.Tables.Count > 0)
			{
				CopyTable(Copy, excelWorksheet);
			}
			if (Copy.PivotTables.Count > 0)
			{
				CopyPivotTable(Copy, excelWorksheet);
			}
			if (Copy.Names.Count > 0)
			{
				CopySheetNames(Copy, excelWorksheet);
			}
			CloneCells(Copy, excelWorksheet);
			if (_pck.Workbook.VbaProject != null)
			{
				string moduleNameFromWorksheet = _pck.Workbook.VbaProject.GetModuleNameFromWorksheet(excelWorksheet);
				_pck.Workbook.VbaProject.Modules.Add(new ExcelVBAModule(excelWorksheet.CodeNameChange)
				{
					Name = moduleNameFromWorksheet,
					Code = Copy.CodeModule.Code,
					Attributes = _pck.Workbook.VbaProject.GetDocumentAttributes(Name, "0{00020820-0000-0000-C000-000000000046}"),
					Type = eModuleType.Document,
					HelpContext = 0
				});
				Copy.CodeModuleName = moduleNameFromWorksheet;
			}
			_worksheets.Add(_worksheets.Count + _pck._worksheetAdd, excelWorksheet);
			XmlNode xmlNode = excelWorksheet.WorksheetXml.SelectSingleNode("//d:pageSetup", _namespaceManager);
			if (xmlNode != null)
			{
				XmlAttribute xmlAttribute = (XmlAttribute)xmlNode.Attributes.GetNamedItem("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
				if (xmlAttribute != null)
				{
					relID = xmlAttribute.Value;
					xmlNode.Attributes.Remove(xmlAttribute);
				}
			}
			return excelWorksheet;
		}
	}

	public ExcelChartsheet AddChart(string Name, eChartType chartType)
	{
		return (ExcelChartsheet)AddSheet(Name, isChart: true, chartType);
	}

	private void CopySheetNames(ExcelWorksheet Copy, ExcelWorksheet added)
	{
		foreach (ExcelNamedRange name in Copy.Names)
		{
			ExcelNamedRange excelNamedRange = (name.IsName ? (string.IsNullOrEmpty(name.NameFormula) ? added.Names.AddValue(name.Name, name.Value) : added.Names.AddFormula(name.Name, name.Formula)) : ((!(name.WorkSheet == Copy.Name)) ? added.Names.Add(name.Name, added.Workbook.Worksheets[name.WorkSheet].Cells[name.FirstAddress]) : added.Names.Add(name.Name, added.Cells[name.FirstAddress])));
			excelNamedRange.NameComment = name.NameComment;
		}
	}

	private void CopyTable(ExcelWorksheet Copy, ExcelWorksheet added)
	{
		string text = "";
		foreach (ExcelTable table in Copy.Tables)
		{
			string outerXml = table.TableXml.OuterXml;
			string text2;
			if (text == "")
			{
				text2 = Copy.Tables.GetNewTableName();
			}
			else
			{
				int num = int.Parse(text.Substring(5)) + 1;
				text2 = $"Table{num}";
				while (_pck.Workbook.ExistsPivotTableName(text2))
				{
					text2 = $"Table{++num}";
				}
			}
			_pck.Workbook.ReadAllTables();
			int id = _pck.Workbook._nextTableID++;
			text = text2;
			XmlDocument xmlDocument = new XmlDocument();
			xmlDocument.LoadXml(outerXml);
			xmlDocument.SelectSingleNode("//d:table/@id", table.NameSpaceManager).Value = id.ToString();
			xmlDocument.SelectSingleNode("//d:table/@name", table.NameSpaceManager).Value = text2;
			xmlDocument.SelectSingleNode("//d:table/@displayName", table.NameSpaceManager).Value = text2;
			outerXml = xmlDocument.OuterXml;
			Uri newUri = XmlHelper.GetNewUri(_pck.Package, "/xl/tables/table{0}.xml", ref id);
			if (_pck.Workbook._nextTableID < id)
			{
				_pck.Workbook._nextTableID = id;
			}
			StreamWriter streamWriter = new StreamWriter(_pck.Package.CreatePart(newUri, "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml", _pck.Compression).GetStream(FileMode.Create, FileAccess.Write));
			streamWriter.Write(outerXml);
			streamWriter.Flush();
			ZipPackageRelationship zipPackageRelationship = added.Part.CreateRelationship(UriHelper.GetRelativeUri(added.WorksheetUri, newUri), TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table");
			if (table.RelationshipID == null)
			{
				XmlNode xmlNode = added.WorksheetXml.SelectSingleNode("//d:tableParts", table.NameSpaceManager);
				if (xmlNode == null)
				{
					added.CreateNode("d:tableParts");
					xmlNode = added.WorksheetXml.SelectSingleNode("//d:tableParts", table.NameSpaceManager);
				}
				XmlElement xmlElement = added.WorksheetXml.CreateElement("tablePart", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
				xmlNode.AppendChild(xmlElement);
				xmlElement.SetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", zipPackageRelationship.Id);
			}
			else
			{
				(added.WorksheetXml.SelectSingleNode($"//d:tableParts/d:tablePart/@r:id[.='{table.RelationshipID}']", table.NameSpaceManager) as XmlAttribute).Value = zipPackageRelationship.Id;
			}
		}
	}

	private void CopyPivotTable(ExcelWorksheet Copy, ExcelWorksheet added)
	{
		string text = "";
		foreach (ExcelPivotTable pivotTable in Copy.PivotTables)
		{
			string outerXml = pivotTable.PivotTableXml.OuterXml;
			string text2;
			if (Copy.Workbook == added.Workbook || added.PivotTables._pivotTableNames.ContainsKey(pivotTable.Name))
			{
				if (text == "")
				{
					text2 = added.PivotTables.GetNewTableName();
				}
				else
				{
					int num = int.Parse(text.Substring(10)) + 1;
					text2 = $"PivotTable{num}";
					while (_pck.Workbook.ExistsPivotTableName(text2))
					{
						text2 = $"PivotTable{++num}";
					}
				}
			}
			else
			{
				text2 = pivotTable.Name;
			}
			text = text2;
			XmlDocument xmlDocument = new XmlDocument();
			xmlDocument.LoadXml(outerXml);
			xmlDocument.SelectSingleNode("//d:pivotTableDefinition/@name", pivotTable.NameSpaceManager).Value = text2;
			int newID = pivotTable.CacheID;
			if (!added.Workbook.ExistsPivotCache(pivotTable.CacheID, ref newID))
			{
				xmlDocument.SelectSingleNode("//d:pivotTableDefinition/@cacheId", pivotTable.NameSpaceManager).Value = newID.ToString();
			}
			outerXml = xmlDocument.OuterXml;
			int id = _pck.Workbook._nextPivotTableID++;
			Uri newUri = XmlHelper.GetNewUri(_pck.Package, "/xl/pivotTables/pivotTable{0}.xml", ref id);
			if (_pck.Workbook._nextPivotTableID < id)
			{
				_pck.Workbook._nextPivotTableID = id;
			}
			ZipPackagePart zipPackagePart = _pck.Package.CreatePart(newUri, "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml", _pck.Compression);
			StreamWriter streamWriter = new StreamWriter(zipPackagePart.GetStream(FileMode.Create, FileAccess.Write));
			streamWriter.Write(outerXml);
			streamWriter.Flush();
			outerXml = pivotTable.CacheDefinition.CacheDefinitionXml.OuterXml;
			Uri newUri2 = XmlHelper.GetNewUri(_pck.Package, "/xl/pivotCache/pivotcachedefinition{0}.xml", ref id);
			ZipPackagePart zipPackagePart2 = _pck.Package.CreatePart(newUri2, "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml", _pck.Compression);
			StreamWriter streamWriter2 = new StreamWriter(zipPackagePart2.GetStream(FileMode.Create, FileAccess.Write));
			streamWriter2.Write(outerXml);
			streamWriter2.Flush();
			added.Workbook.AddPivotTable(newID.ToString(), newUri2);
			outerXml = "<pivotCacheRecords xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" count=\"0\" />";
			Uri uri = new Uri($"/xl/pivotCache/pivotCacheRecords{id}.xml", UriKind.Relative);
			while (_pck.Package.PartExists(uri))
			{
				uri = new Uri($"/xl/pivotCache/pivotCacheRecords{++id}.xml", UriKind.Relative);
			}
			StreamWriter streamWriter3 = new StreamWriter(_pck.Package.CreatePart(uri, "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml", _pck.Compression).GetStream(FileMode.Create, FileAccess.Write));
			streamWriter3.Write(outerXml);
			streamWriter3.Flush();
			added.Part.CreateRelationship(UriHelper.ResolvePartUri(added.WorksheetUri, newUri), TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable");
			zipPackagePart.CreateRelationship(UriHelper.ResolvePartUri(pivotTable.Relationship.SourceUri, newUri2), pivotTable.CacheDefinition.Relationship.TargetMode, pivotTable.CacheDefinition.Relationship.RelationshipType);
			zipPackagePart2.CreateRelationship(UriHelper.ResolvePartUri(newUri2, uri), TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords");
		}
		added._pivotTables = null;
	}

	private void CopyHeaderFooterPictures(ExcelWorksheet Copy, ExcelWorksheet added)
	{
		if (Copy.TopNode != null && Copy.TopNode.SelectSingleNode("d:headerFooter", base.NameSpaceManager) == null)
		{
			return;
		}
		if (Copy.HeaderFooter._oddHeader != null)
		{
			CopyText(Copy.HeaderFooter._oddHeader, added.HeaderFooter.OddHeader);
		}
		if (Copy.HeaderFooter._oddFooter != null)
		{
			CopyText(Copy.HeaderFooter._oddFooter, added.HeaderFooter.OddFooter);
		}
		if (Copy.HeaderFooter._evenHeader != null)
		{
			CopyText(Copy.HeaderFooter._evenHeader, added.HeaderFooter.EvenHeader);
		}
		if (Copy.HeaderFooter._evenFooter != null)
		{
			CopyText(Copy.HeaderFooter._evenFooter, added.HeaderFooter.EvenFooter);
		}
		if (Copy.HeaderFooter._firstHeader != null)
		{
			CopyText(Copy.HeaderFooter._firstHeader, added.HeaderFooter.FirstHeader);
		}
		if (Copy.HeaderFooter._firstFooter != null)
		{
			CopyText(Copy.HeaderFooter._firstFooter, added.HeaderFooter.FirstFooter);
		}
		if (Copy.HeaderFooter.Pictures.Count <= 0)
		{
			return;
		}
		_ = Copy.HeaderFooter.Pictures.Uri;
		XmlHelper.GetNewUri(_pck.Package, "/xl/drawings/vmlDrawing{0}.vml");
		added.DeleteNode("d:legacyDrawingHF");
		foreach (ExcelVmlDrawingPicture item in (IEnumerable)Copy.HeaderFooter.Pictures)
		{
			ExcelVmlDrawingPicture excelVmlDrawingPicture2 = added.HeaderFooter.Pictures.Add(item.Id, item.ImageUri, item.Title, item.Width, item.Height);
			foreach (XmlAttribute attribute in item.TopNode.Attributes)
			{
				(excelVmlDrawingPicture2.TopNode as XmlElement).SetAttribute(attribute.Name, attribute.Value);
			}
			excelVmlDrawingPicture2.TopNode.InnerXml = item.TopNode.InnerXml;
		}
	}

	private void CopyText(ExcelHeaderFooterText from, ExcelHeaderFooterText to)
	{
		to.LeftAlignedText = from.LeftAlignedText;
		to.CenteredText = from.CenteredText;
		to.RightAlignedText = from.RightAlignedText;
	}

	private void CloneCells(ExcelWorksheet Copy, ExcelWorksheet added)
	{
		bool flag = Copy.Workbook == _pck.Workbook;
		bool doAdjustDrawings = _pck.DoAdjustDrawings;
		_pck.DoAdjustDrawings = false;
		foreach (string mergedCell in Copy.MergedCells)
		{
			added.MergedCells.Add(new ExcelAddress(mergedCell), doValidate: false);
		}
		foreach (int key in Copy._sharedFormulas.Keys)
		{
			added._sharedFormulas.Add(key, Copy._sharedFormulas[key].Clone());
		}
		Dictionary<int, int> dictionary = new Dictionary<int, int>();
		CellsStoreEnumerator<ExcelCoreValue> cellsStoreEnumerator = new CellsStoreEnumerator<ExcelCoreValue>(Copy._values);
		while (cellsStoreEnumerator.Next())
		{
			int row = cellsStoreEnumerator.Row;
			int column = cellsStoreEnumerator.Column;
			int num = 0;
			if (row == 0)
			{
				if (Copy.GetValueInner(row, column) is ExcelColumn excelColumn)
				{
					ExcelColumn excelColumn2 = excelColumn.Clone(added, excelColumn.ColumnMin);
					excelColumn2.StyleID = excelColumn.StyleID;
					added.SetValueInner(row, column, excelColumn2);
					num = excelColumn.StyleID;
				}
			}
			else if (column == 0)
			{
				ExcelRow excelRow = Copy.Row(row);
				if (excelRow != null)
				{
					excelRow.Clone(added);
					num = excelRow.StyleID;
				}
			}
			else
			{
				num = CopyValues(Copy, added, row, column);
			}
			if (!flag)
			{
				if (dictionary.ContainsKey(num))
				{
					added.SetStyleInner(row, column, dictionary[num]);
					continue;
				}
				int num2 = added.Workbook.Styles.CloneStyle(Copy.Workbook.Styles, num);
				dictionary.Add(num, num2);
				added.SetStyleInner(row, column, num2);
			}
		}
		added._package.DoAdjustDrawings = doAdjustDrawings;
	}

	private int CopyValues(ExcelWorksheet Copy, ExcelWorksheet added, int row, int col)
	{
		added.SetValueInner(row, col, Copy.GetValueInner(row, col));
		byte value = 0;
		if (Copy._flags.Exists(row, col, ref value))
		{
			added._flags.SetValue(row, col, value);
		}
		object value2 = Copy._formulas.GetValue(row, col);
		if (value2 != null)
		{
			added.SetFormula(row, col, value2);
		}
		int styleInner = Copy.GetStyleInner(row, col);
		if (styleInner != 0)
		{
			added.SetStyleInner(row, col, styleInner);
		}
		object value3 = Copy._formulas.GetValue(row, col);
		if (value3 != null)
		{
			added._formulas.SetValue(row, col, value3);
		}
		return styleInner;
	}

	private void CopyComment(ExcelWorksheet Copy, ExcelWorksheet workSheet)
	{
		string innerXml = Copy.Comments.CommentXml.InnerXml;
		Uri uri = new Uri($"/xl/comments{workSheet.SheetID}.xml", UriKind.Relative);
		if (_pck.Package.PartExists(uri))
		{
			uri = XmlHelper.GetNewUri(_pck.Package, "/xl/drawings/vmldrawing{0}.vml");
		}
		StreamWriter streamWriter = new StreamWriter(_pck.Package.CreatePart(uri, "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml", _pck.Compression).GetStream(FileMode.Create, FileAccess.Write));
		streamWriter.Write(innerXml);
		streamWriter.Flush();
		workSheet.Part.CreateRelationship(UriHelper.GetRelativeUri(workSheet.WorksheetUri, uri), TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments");
		innerXml = Copy.VmlDrawingsComments.VmlDrawingXml.InnerXml;
		Uri uri2 = new Uri($"/xl/drawings/vmldrawing{workSheet.SheetID}.vml", UriKind.Relative);
		if (_pck.Package.PartExists(uri2))
		{
			uri2 = XmlHelper.GetNewUri(_pck.Package, "/xl/drawings/vmldrawing{0}.vml");
		}
		StreamWriter streamWriter2 = new StreamWriter(_pck.Package.CreatePart(uri2, "application/vnd.openxmlformats-officedocument.vmlDrawing", _pck.Compression).GetStream(FileMode.Create, FileAccess.Write));
		streamWriter2.Write(innerXml);
		streamWriter2.Flush();
		ZipPackageRelationship zipPackageRelationship = workSheet.Part.CreateRelationship(UriHelper.GetRelativeUri(workSheet.WorksheetUri, uri2), TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing");
		XmlElement xmlElement = workSheet.WorksheetXml.SelectSingleNode("//d:legacyDrawing", _namespaceManager) as XmlElement;
		if (xmlElement == null)
		{
			workSheet.CreateNode("d:legacyDrawing");
			xmlElement = workSheet.WorksheetXml.SelectSingleNode("//d:legacyDrawing", _namespaceManager) as XmlElement;
		}
		xmlElement.SetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", zipPackageRelationship.Id);
	}

	private void CopyDrawing(ExcelWorksheet Copy, ExcelWorksheet workSheet)
	{
		string outerXml = Copy.Drawings.DrawingXml.OuterXml;
		Uri uri = new Uri($"/xl/drawings/drawing{workSheet.SheetID}.xml", UriKind.Relative);
		ZipPackagePart zipPackagePart = _pck.Package.CreatePart(uri, "application/vnd.openxmlformats-officedocument.drawing+xml", _pck.Compression);
		StreamWriter streamWriter = new StreamWriter(zipPackagePart.GetStream(FileMode.Create, FileAccess.Write));
		streamWriter.Write(outerXml);
		streamWriter.Flush();
		XmlDocument xmlDocument = new XmlDocument();
		xmlDocument.LoadXml(outerXml);
		ZipPackageRelationship zipPackageRelationship = workSheet.Part.CreateRelationship(UriHelper.GetRelativeUri(workSheet.WorksheetUri, uri), TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing");
		(workSheet.WorksheetXml.SelectSingleNode("//d:drawing", _namespaceManager) as XmlElement).SetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", zipPackageRelationship.Id);
		for (int i = 0; i < Copy.Drawings.Count; i++)
		{
			ExcelDrawing excelDrawing = Copy.Drawings[i];
			excelDrawing.AdjustPositionAndSize();
			if (excelDrawing is ExcelChart)
			{
				outerXml = (excelDrawing as ExcelChart).ChartXml.InnerXml;
				Uri newUri = XmlHelper.GetNewUri(_pck.Package, "/xl/charts/chart{0}.xml");
				StreamWriter streamWriter2 = new StreamWriter(_pck.Package.CreatePart(newUri, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml", _pck.Compression).GetStream(FileMode.Create, FileAccess.Write));
				streamWriter2.Write(outerXml);
				streamWriter2.Flush();
				string value = excelDrawing.TopNode.SelectSingleNode("xdr:graphicFrame/a:graphic/a:graphicData/c:chart/@r:id", Copy.Drawings.NameSpaceManager).Value;
				ZipPackageRelationship zipPackageRelationship2 = zipPackagePart.CreateRelationship(UriHelper.GetRelativeUri(uri, newUri), TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart");
				(xmlDocument.SelectSingleNode($"//c:chart/@r:id[.='{value}']", Copy.Drawings.NameSpaceManager) as XmlAttribute).Value = zipPackageRelationship2.Id;
			}
			else if (excelDrawing is ExcelPicture)
			{
				ExcelPicture excelPicture = excelDrawing as ExcelPicture;
				Uri uriPic = excelPicture.UriPic;
				if (!workSheet.Workbook._package.Package.PartExists(uriPic))
				{
					ZipPackagePart zipPackagePart2 = workSheet.Workbook._package.Package.CreatePart(uriPic, excelPicture.ContentType, CompressionLevel.Level0);
					excelPicture.Image.Save(zipPackagePart2.GetStream(FileMode.Create, FileAccess.Write), ExcelPicture.GetImageFormat(excelPicture.ContentType));
				}
				ZipPackageRelationship zipPackageRelationship3 = zipPackagePart.CreateRelationship(UriHelper.GetRelativeUri(workSheet.WorksheetUri, uriPic), TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
				XmlNode xmlNode = xmlDocument.SelectSingleNode($"//xdr:pic/xdr:nvPicPr/xdr:cNvPr/@name[.='{excelPicture.Name}']/../../../xdr:blipFill/a:blip/@r:embed", Copy.Drawings.NameSpaceManager);
				if (xmlNode != null)
				{
					xmlNode.Value = zipPackageRelationship3.Id;
				}
				if (_pck._images.ContainsKey(excelPicture.ImageHash))
				{
					_pck._images[excelPicture.ImageHash].RefCount++;
				}
			}
		}
		StreamWriter streamWriter3 = new StreamWriter(zipPackagePart.GetStream(FileMode.Create, FileAccess.Write));
		streamWriter3.Write(xmlDocument.OuterXml);
		streamWriter3.Flush();
		for (int j = 0; j < Copy.Drawings.Count; j++)
		{
			ExcelDrawing excelDrawing2 = Copy.Drawings[j];
			ExcelDrawing excelDrawing3 = workSheet.Drawings[j];
			if (excelDrawing3 != null)
			{
				excelDrawing3._left = excelDrawing2._left;
				excelDrawing3._top = excelDrawing2._top;
				excelDrawing3._height = excelDrawing2._height;
				excelDrawing3._width = excelDrawing2._width;
			}
		}
	}

	private void CopyVmlDrawing(ExcelWorksheet origSheet, ExcelWorksheet newSheet)
	{
		string outerXml = origSheet.VmlDrawingsComments.VmlDrawingXml.OuterXml;
		Uri uri = new Uri($"/xl/drawings/vmlDrawing{newSheet.SheetID}.vml", UriKind.Relative);
		using (StreamWriter streamWriter = new StreamWriter(_pck.Package.CreatePart(uri, "application/vnd.openxmlformats-officedocument.vmlDrawing", _pck.Compression).GetStream(FileMode.Create, FileAccess.Write)))
		{
			streamWriter.Write(outerXml);
			streamWriter.Flush();
		}
		ZipPackageRelationship zipPackageRelationship = newSheet.Part.CreateRelationship(UriHelper.GetRelativeUri(newSheet.WorksheetUri, uri), TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing");
		XmlElement xmlElement = newSheet.WorksheetXml.SelectSingleNode("//d:legacyDrawing", _namespaceManager) as XmlElement;
		if (xmlElement == null)
		{
			xmlElement = newSheet.WorksheetXml.CreateNode(XmlNodeType.Entity, "//d:legacyDrawing", _namespaceManager.LookupNamespace("d")) as XmlElement;
		}
		xmlElement?.SetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", zipPackageRelationship.Id);
	}

	private string CreateWorkbookRel(string Name, int sheetID, Uri uriWorksheet, bool isChart)
	{
		ZipPackageRelationship zipPackageRelationship = _pck.Workbook.Part.CreateRelationship(UriHelper.GetRelativeUri(_pck.Workbook.WorkbookUri, uriWorksheet), TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/" + (isChart ? "chartsheet" : "worksheet"));
		_pck.Package.Flush();
		XmlElement xmlElement = _pck.Workbook.WorkbookXml.CreateElement("sheet", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
		xmlElement.SetAttribute("name", Name);
		xmlElement.SetAttribute("sheetId", sheetID.ToString());
		xmlElement.SetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", zipPackageRelationship.Id);
		base.TopNode.AppendChild(xmlElement);
		return zipPackageRelationship.Id;
	}

	private void GetSheetURI(ref string Name, out int sheetID, out Uri uriWorksheet, bool isChart)
	{
		Name = ValidateFixSheetName(Name);
		sheetID = ((!this.Any()) ? 1 : (this.Max((ExcelWorksheet ws) => ws.SheetID) + 1));
		int num = sheetID;
		do
		{
			if (isChart)
			{
				uriWorksheet = new Uri("/xl/chartsheets/chartsheet" + num + ".xml", UriKind.Relative);
			}
			else
			{
				uriWorksheet = new Uri("/xl/worksheets/sheet" + num + ".xml", UriKind.Relative);
			}
			num++;
		}
		while (_pck.Package.PartExists(uriWorksheet));
	}

	internal string ValidateFixSheetName(string Name)
	{
		if (ValidateName(Name))
		{
			if (Name.IndexOf(':') > -1)
			{
				Name = Name.Replace(":", " ");
			}
			if (Name.IndexOf('/') > -1)
			{
				Name = Name.Replace("/", " ");
			}
			if (Name.IndexOf('\\') > -1)
			{
				Name = Name.Replace("\\", " ");
			}
			if (Name.IndexOf('?') > -1)
			{
				Name = Name.Replace("?", " ");
			}
			if (Name.IndexOf('[') > -1)
			{
				Name = Name.Replace("[", " ");
			}
			if (Name.IndexOf(']') > -1)
			{
				Name = Name.Replace("]", " ");
			}
		}
		if (Name.Trim() == "")
		{
			throw new ArgumentException("The worksheet can not have an empty name");
		}
		if (Name.StartsWith("'") || Name.EndsWith("'"))
		{
			throw new ArgumentException("The worksheet name can not start or end with an apostrophe.");
		}
		if (Name.Length > 31)
		{
			Name = Name.Substring(0, 31);
		}
		return Name;
	}

	private bool ValidateName(string Name)
	{
		return Regex.IsMatch(Name, ":|\\?|/|\\\\|\\[|\\]");
	}

	internal XmlDocument CreateNewWorksheet(bool isChart)
	{
		XmlDocument xmlDocument = new XmlDocument();
		XmlElement xmlElement = xmlDocument.CreateElement(isChart ? "chartsheet" : "worksheet", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
		xmlElement.SetAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
		xmlDocument.AppendChild(xmlElement);
		if (isChart)
		{
			XmlElement newChild = xmlDocument.CreateElement("sheetPr", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
			xmlElement.AppendChild(newChild);
			XmlElement xmlElement2 = xmlDocument.CreateElement("sheetViews", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
			xmlElement.AppendChild(xmlElement2);
			XmlElement xmlElement3 = xmlDocument.CreateElement("sheetView", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
			xmlElement3.SetAttribute("workbookViewId", "0");
			xmlElement3.SetAttribute("zoomToFit", "1");
			xmlElement2.AppendChild(xmlElement3);
		}
		else
		{
			XmlElement xmlElement4 = xmlDocument.CreateElement("sheetViews", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
			xmlElement.AppendChild(xmlElement4);
			XmlElement xmlElement5 = xmlDocument.CreateElement("sheetView", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
			xmlElement5.SetAttribute("workbookViewId", "0");
			xmlElement4.AppendChild(xmlElement5);
			XmlElement xmlElement6 = xmlDocument.CreateElement("sheetFormatPr", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
			xmlElement6.SetAttribute("defaultRowHeight", "15");
			xmlElement.AppendChild(xmlElement6);
			XmlElement newChild2 = xmlDocument.CreateElement("sheetData", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
			xmlElement.AppendChild(newChild2);
		}
		return xmlDocument;
	}

	public void Delete(int Index)
	{
		foreach (KeyValuePair<int, ExcelWorksheet> worksheet in _worksheets)
		{
			_ = worksheet.Value.Drawings;
		}
		ExcelWorksheet excelWorksheet = _worksheets[Index];
		if (excelWorksheet.Drawings.Count > 0)
		{
			excelWorksheet.Drawings.ClearDrawings();
		}
		if (!(excelWorksheet is ExcelChartsheet) && excelWorksheet.Comments.Count > 0)
		{
			excelWorksheet.Comments.Clear();
		}
		DeleteRelationsAndParts(excelWorksheet.Part);
		_pck.Workbook.Part.DeleteRelationship(excelWorksheet.RelationshipID);
		XmlNode xmlNode = _pck.Workbook.WorkbookXml.SelectSingleNode("//d:workbook/d:sheets", _namespaceManager);
		if (xmlNode != null)
		{
			XmlNode xmlNode2 = xmlNode.SelectSingleNode($"./d:sheet[@sheetId={excelWorksheet.SheetID}]", _namespaceManager);
			if (xmlNode2 != null)
			{
				xmlNode.RemoveChild(xmlNode2);
			}
		}
		_worksheets.Remove(Index);
		if (_pck.Workbook.VbaProject != null)
		{
			_pck.Workbook.VbaProject.Modules.Remove(excelWorksheet.CodeModule);
		}
		ReindexWorksheetDictionary();
		if (_pck.Workbook.View.ActiveTab >= _pck.Workbook.Worksheets.Count)
		{
			_pck.Workbook.View.ActiveTab = _pck.Workbook.View.ActiveTab - 1;
		}
		if (_pck.Workbook.View.ActiveTab == excelWorksheet.SheetID)
		{
			_pck.Workbook.Worksheets[_pck._worksheetAdd].View.TabSelected = true;
		}
		excelWorksheet = null;
	}

	private void DeleteRelationsAndParts(ZipPackagePart part)
	{
		List<ZipPackageRelationship> list = part.GetRelationships().ToList();
		for (int i = 0; i < list.Count; i++)
		{
			ZipPackageRelationship zipPackageRelationship = list[i];
			if (zipPackageRelationship.RelationshipType != "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
			{
				DeleteRelationsAndParts(_pck.Package.GetPart(UriHelper.ResolvePartUri(zipPackageRelationship.SourceUri, zipPackageRelationship.TargetUri)));
			}
			part.DeleteRelationship(zipPackageRelationship.Id);
		}
		_pck.Package.DeletePart(part.Uri);
	}

	public void Delete(string name)
	{
		ExcelWorksheet excelWorksheet = this[name];
		if (excelWorksheet == null)
		{
			throw new ArgumentException($"Could not find worksheet to delete '{name}'");
		}
		Delete(excelWorksheet.PositionID);
	}

	public void Delete(ExcelWorksheet Worksheet)
	{
		if (Worksheet.PositionID <= _worksheets.Count && Worksheet == _worksheets[Worksheet.PositionID])
		{
			Delete(Worksheet.PositionID);
			return;
		}
		throw new ArgumentException("Worksheet is not in the collection.");
	}

	internal void ReindexWorksheetDictionary()
	{
		int worksheetAdd = _pck._worksheetAdd;
		Dictionary<int, ExcelWorksheet> dictionary = new Dictionary<int, ExcelWorksheet>();
		foreach (KeyValuePair<int, ExcelWorksheet> worksheet in _worksheets)
		{
			worksheet.Value.PositionID = worksheetAdd;
			dictionary.Add(worksheetAdd++, worksheet.Value);
		}
		_worksheets = dictionary;
	}

	public ExcelWorksheet Copy(string Name, string NewName)
	{
		ExcelWorksheet excelWorksheet = this[Name];
		if (excelWorksheet == null)
		{
			throw new ArgumentException($"Copy worksheet error: Could not find worksheet to copy '{Name}'");
		}
		return Add(NewName, excelWorksheet);
	}

	internal ExcelWorksheet GetBySheetID(int localSheetID)
	{
		using (IEnumerator<ExcelWorksheet> enumerator = GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				ExcelWorksheet current = enumerator.Current;
				if (current.SheetID == localSheetID)
				{
					return current;
				}
			}
		}
		return null;
	}

	private ExcelWorksheet GetByName(string Name)
	{
		if (string.IsNullOrEmpty(Name))
		{
			return null;
		}
		ExcelWorksheet result = null;
		foreach (ExcelWorksheet value in _worksheets.Values)
		{
			if (value.Name.Equals(Name, StringComparison.OrdinalIgnoreCase))
			{
				result = value;
			}
		}
		return result;
	}

	public void MoveBefore(string sourceName, string targetName)
	{
		Move(sourceName, targetName, placeAfter: false);
	}

	public void MoveBefore(int sourcePositionId, int targetPositionId)
	{
		Move(sourcePositionId, targetPositionId, placeAfter: false);
	}

	public void MoveAfter(string sourceName, string targetName)
	{
		Move(sourceName, targetName, placeAfter: true);
	}

	public void MoveAfter(int sourcePositionId, int targetPositionId)
	{
		Move(sourcePositionId, targetPositionId, placeAfter: true);
	}

	public void MoveToStart(string sourceName)
	{
		ExcelWorksheet excelWorksheet = this[sourceName];
		if (excelWorksheet == null)
		{
			throw new Exception($"Move worksheet error: Could not find worksheet to move '{sourceName}'");
		}
		Move(excelWorksheet.PositionID, _pck._worksheetAdd, placeAfter: false);
	}

	public void MoveToStart(int sourcePositionId)
	{
		Move(sourcePositionId, _pck._worksheetAdd, placeAfter: false);
	}

	public void MoveToEnd(string sourceName)
	{
		ExcelWorksheet excelWorksheet = this[sourceName];
		if (excelWorksheet == null)
		{
			throw new Exception($"Move worksheet error: Could not find worksheet to move '{sourceName}'");
		}
		Move(excelWorksheet.PositionID, _worksheets.Count + (_pck._worksheetAdd - 1), placeAfter: true);
	}

	public void MoveToEnd(int sourcePositionId)
	{
		Move(sourcePositionId, _worksheets.Count + (_pck._worksheetAdd - 1), placeAfter: true);
	}

	private void Move(string sourceName, string targetName, bool placeAfter)
	{
		ExcelWorksheet excelWorksheet = this[sourceName];
		if (excelWorksheet == null)
		{
			throw new Exception($"Move worksheet error: Could not find worksheet to move '{sourceName}'");
		}
		ExcelWorksheet excelWorksheet2 = this[targetName];
		if (excelWorksheet2 == null)
		{
			throw new Exception($"Move worksheet error: Could not find worksheet to move '{targetName}'");
		}
		Move(excelWorksheet.PositionID, excelWorksheet2.PositionID, placeAfter);
	}

	private void Move(int sourcePositionId, int targetPositionId, bool placeAfter)
	{
		if (sourcePositionId == targetPositionId)
		{
			return;
		}
		lock (_worksheets)
		{
			ExcelWorksheet excelWorksheet = this[sourcePositionId];
			if (excelWorksheet == null)
			{
				throw new Exception($"Move worksheet error: Could not find worksheet at position '{sourcePositionId}'");
			}
			ExcelWorksheet excelWorksheet2 = this[targetPositionId];
			if (excelWorksheet2 == null)
			{
				throw new Exception($"Move worksheet error: Could not find worksheet at position '{targetPositionId}'");
			}
			if (sourcePositionId == targetPositionId && _worksheets.Count < 2)
			{
				return;
			}
			int worksheetAdd = _pck._worksheetAdd;
			Dictionary<int, ExcelWorksheet> dictionary = new Dictionary<int, ExcelWorksheet>();
			foreach (KeyValuePair<int, ExcelWorksheet> worksheet in _worksheets)
			{
				if (worksheet.Key == targetPositionId)
				{
					if (!placeAfter)
					{
						excelWorksheet.PositionID = worksheetAdd;
						dictionary.Add(worksheetAdd++, excelWorksheet);
					}
					worksheet.Value.PositionID = worksheetAdd;
					dictionary.Add(worksheetAdd++, worksheet.Value);
					if (placeAfter)
					{
						excelWorksheet.PositionID = worksheetAdd;
						dictionary.Add(worksheetAdd++, excelWorksheet);
					}
				}
				else if (worksheet.Key != sourcePositionId)
				{
					worksheet.Value.PositionID = worksheetAdd;
					dictionary.Add(worksheetAdd++, worksheet.Value);
				}
			}
			_worksheets = dictionary;
			MoveSheetXmlNode(excelWorksheet, excelWorksheet2, placeAfter);
		}
	}

	private void MoveSheetXmlNode(ExcelWorksheet sourceSheet, ExcelWorksheet targetSheet, bool placeAfter)
	{
		lock (base.TopNode.OwnerDocument)
		{
			XmlNode xmlNode = base.TopNode.SelectSingleNode($"d:sheet[@sheetId = '{sourceSheet.SheetID}']", _namespaceManager);
			XmlNode xmlNode2 = base.TopNode.SelectSingleNode($"d:sheet[@sheetId = '{targetSheet.SheetID}']", _namespaceManager);
			if (xmlNode == null || xmlNode2 == null)
			{
				throw new Exception("Source SheetId and Target SheetId must be valid");
			}
			if (placeAfter)
			{
				base.TopNode.InsertAfter(xmlNode, xmlNode2);
			}
			else
			{
				base.TopNode.InsertBefore(xmlNode, xmlNode2);
			}
		}
	}

	public void Dispose()
	{
		if (_worksheets == null)
		{
			return;
		}
		foreach (ExcelWorksheet value in _worksheets.Values)
		{
			((IDisposable)value).Dispose();
		}
		_worksheets = null;
		_pck = null;
	}
}
