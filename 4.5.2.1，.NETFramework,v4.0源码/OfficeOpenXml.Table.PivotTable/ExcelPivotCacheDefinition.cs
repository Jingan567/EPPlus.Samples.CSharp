using System;
using System.Linq;
using System.Text;
using System.Xml;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Table.PivotTable;

public class ExcelPivotCacheDefinition : XmlHelper
{
	private const string _sourceWorksheetPath = "d:cacheSource/d:worksheetSource/@sheet";

	internal const string _sourceNamePath = "d:cacheSource/d:worksheetSource/@name";

	internal const string _sourceAddressPath = "d:cacheSource/d:worksheetSource/@ref";

	internal ExcelRangeBase _sourceRange;

	internal ZipPackagePart Part { get; set; }

	public XmlDocument CacheDefinitionXml { get; private set; }

	public Uri CacheDefinitionUri { get; internal set; }

	internal Uri CacheRecordUri { get; set; }

	internal ZipPackageRelationship Relationship { get; set; }

	internal ZipPackageRelationship RecordRelationship { get; set; }

	internal string RecordRelationshipID
	{
		get
		{
			return GetXmlNodeString("@r:id");
		}
		set
		{
			SetXmlNodeString("@r:id", value);
		}
	}

	public ExcelPivotTable PivotTable { get; private set; }

	public ExcelRangeBase SourceRange
	{
		get
		{
			if (_sourceRange == null)
			{
				if (CacheSource != eSourceType.Worksheet)
				{
					throw new ArgumentException("The cachesource is not a worksheet");
				}
				ExcelWorksheet excelWorksheet = PivotTable.WorkSheet.Workbook.Worksheets[GetXmlNodeString("d:cacheSource/d:worksheetSource/@sheet")];
				if (excelWorksheet == null)
				{
					string xmlNodeString = GetXmlNodeString("d:cacheSource/d:worksheetSource/@name");
					foreach (ExcelNamedRange name in PivotTable.WorkSheet.Workbook.Names)
					{
						if (xmlNodeString.Equals(name.Name, StringComparison.OrdinalIgnoreCase))
						{
							_sourceRange = name;
							return _sourceRange;
						}
					}
					foreach (ExcelWorksheet worksheet in PivotTable.WorkSheet.Workbook.Worksheets)
					{
						if (worksheet.Tables._tableNames.ContainsKey(xmlNodeString))
						{
							_sourceRange = worksheet.Cells[worksheet.Tables[xmlNodeString].Address.Address];
							break;
						}
						foreach (ExcelNamedRange name2 in worksheet.Names)
						{
							if (xmlNodeString.Equals(name2.Name, StringComparison.OrdinalIgnoreCase))
							{
								_sourceRange = name2;
								break;
							}
						}
					}
				}
				else
				{
					_sourceRange = excelWorksheet.Cells[GetXmlNodeString("d:cacheSource/d:worksheetSource/@ref")];
				}
			}
			return _sourceRange;
		}
		set
		{
			if (PivotTable.WorkSheet.Workbook != value.Worksheet.Workbook)
			{
				throw new ArgumentException("Range must be in the same package as the pivottable");
			}
			ExcelRangeBase sourceRange = SourceRange;
			if (value.End.Column - value.Start.Column != sourceRange.End.Column - sourceRange.Start.Column)
			{
				throw new ArgumentException("Can not change the number of columns(fields) in the SourceRange");
			}
			SetXmlNodeString("d:cacheSource/d:worksheetSource/@sheet", value.Worksheet.Name);
			SetXmlNodeString("d:cacheSource/d:worksheetSource/@ref", value.FirstAddress);
			_sourceRange = value;
		}
	}

	public eSourceType CacheSource
	{
		get
		{
			string xmlNodeString = GetXmlNodeString("d:cacheSource/@type");
			if (xmlNodeString == "")
			{
				return eSourceType.Worksheet;
			}
			return (eSourceType)Enum.Parse(typeof(eSourceType), xmlNodeString, ignoreCase: true);
		}
	}

	internal ExcelPivotCacheDefinition(XmlNamespaceManager ns, ExcelPivotTable pivotTable)
		: base(ns, null)
	{
		foreach (ZipPackageRelationship item in pivotTable.Part.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition"))
		{
			Relationship = item;
		}
		CacheDefinitionUri = UriHelper.ResolvePartUri(Relationship.SourceUri, Relationship.TargetUri);
		ZipPackage package = pivotTable.WorkSheet._package.Package;
		Part = package.GetPart(CacheDefinitionUri);
		CacheDefinitionXml = new XmlDocument();
		XmlHelper.LoadXmlSafe(CacheDefinitionXml, Part.GetStream());
		base.TopNode = CacheDefinitionXml.DocumentElement;
		PivotTable = pivotTable;
		if (CacheSource == eSourceType.Worksheet)
		{
			string worksheetName = GetXmlNodeString("d:cacheSource/d:worksheetSource/@sheet");
			if (pivotTable.WorkSheet.Workbook.Worksheets.Any((ExcelWorksheet t) => t.Name == worksheetName))
			{
				_sourceRange = pivotTable.WorkSheet.Workbook.Worksheets[worksheetName].Cells[GetXmlNodeString("d:cacheSource/d:worksheetSource/@ref")];
			}
		}
	}

	internal ExcelPivotCacheDefinition(XmlNamespaceManager ns, ExcelPivotTable pivotTable, ExcelRangeBase sourceAddress, int tblId)
		: base(ns, null)
	{
		PivotTable = pivotTable;
		ZipPackage package = pivotTable.WorkSheet._package.Package;
		CacheDefinitionXml = new XmlDocument();
		XmlHelper.LoadXmlSafe(CacheDefinitionXml, GetStartXml(sourceAddress), Encoding.UTF8);
		CacheDefinitionUri = XmlHelper.GetNewUri(package, "/xl/pivotCache/pivotCacheDefinition{0}.xml", ref tblId);
		Part = package.CreatePart(CacheDefinitionUri, "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml");
		base.TopNode = CacheDefinitionXml.DocumentElement;
		CacheRecordUri = XmlHelper.GetNewUri(package, "/xl/pivotCache/pivotCacheRecords{0}.xml", ref tblId);
		XmlDocument xmlDocument = new XmlDocument();
		xmlDocument.LoadXml("<pivotCacheRecords xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" count=\"0\" />");
		ZipPackagePart zipPackagePart = package.CreatePart(CacheRecordUri, "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml");
		xmlDocument.Save(zipPackagePart.GetStream());
		RecordRelationship = Part.CreateRelationship(UriHelper.ResolvePartUri(CacheDefinitionUri, CacheRecordUri), TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords");
		RecordRelationshipID = RecordRelationship.Id;
		CacheDefinitionXml.Save(Part.GetStream());
	}

	private string GetStartXml(ExcelRangeBase sourceAddress)
	{
		string text = "<pivotCacheDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"\" refreshOnLoad=\"1\" refreshedBy=\"SomeUser\" refreshedDate=\"40504.582403125001\" createdVersion=\"1\" refreshedVersion=\"3\" recordCount=\"5\" upgradeOnRefresh=\"1\">";
		text += "<cacheSource type=\"worksheet\">";
		text += $"<worksheetSource ref=\"{sourceAddress.Address}\" sheet=\"{sourceAddress.WorkSheet}\" /> ";
		text += "</cacheSource>";
		text += $"<cacheFields count=\"{sourceAddress._toCol - sourceAddress._fromCol + 1}\">";
		ExcelWorksheet excelWorksheet = PivotTable.WorkSheet.Workbook.Worksheets[sourceAddress.WorkSheet];
		for (int i = sourceAddress._fromCol; i <= sourceAddress._toCol; i++)
		{
			text = ((excelWorksheet != null && excelWorksheet.GetValueInner(sourceAddress._fromRow, i) != null && !(excelWorksheet.GetValueInner(sourceAddress._fromRow, i).ToString().Trim() == "")) ? (text + $"<cacheField name=\"{excelWorksheet.GetValueInner(sourceAddress._fromRow, i)}\" numFmtId=\"0\">") : (text + $"<cacheField name=\"Column{i - sourceAddress._fromCol + 1}\" numFmtId=\"0\">"));
			text += "<sharedItems containsBlank=\"1\" /> ";
			text += "</cacheField>";
		}
		text += "</cacheFields>";
		return text + "</pivotCacheDefinition>";
	}
}
