using System;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Table.PivotTable;

public class ExcelPivotTable : XmlHelper
{
	private const string NAME_PATH = "@name";

	private const string DISPLAY_NAME_PATH = "@displayName";

	private ExcelPivotCacheDefinition _cacheDefinition;

	private const string FIRSTHEADERROW_PATH = "d:location/@firstHeaderRow";

	private const string FIRSTDATAROW_PATH = "d:location/@firstDataRow";

	private const string FIRSTDATACOL_PATH = "d:location/@firstDataCol";

	private ExcelPivotTableFieldCollection _fields;

	private ExcelPivotTableRowColumnFieldCollection _rowFields;

	private ExcelPivotTableRowColumnFieldCollection _columnFields;

	private ExcelPivotTableDataFieldCollection _dataFields;

	private ExcelPivotTableRowColumnFieldCollection _pageFields;

	private const string STYLENAME_PATH = "d:pivotTableStyleInfo/@name";

	private TableStyles _tableStyle = TableStyles.Medium6;

	internal ZipPackagePart Part { get; set; }

	public XmlDocument PivotTableXml { get; private set; }

	public Uri PivotTableUri { get; internal set; }

	internal ZipPackageRelationship Relationship { get; set; }

	public string Name
	{
		get
		{
			return GetXmlNodeString("@name");
		}
		set
		{
			if (WorkSheet.Workbook.ExistsTableName(value))
			{
				throw new ArgumentException("PivotTable name is not unique");
			}
			string name = Name;
			if (WorkSheet.Tables._tableNames.ContainsKey(name))
			{
				int value2 = WorkSheet.Tables._tableNames[name];
				WorkSheet.Tables._tableNames.Remove(name);
				WorkSheet.Tables._tableNames.Add(value, value2);
			}
			SetXmlNodeString("@name", value);
			SetXmlNodeString("@displayName", cleanDisplayName(value));
		}
	}

	public ExcelPivotCacheDefinition CacheDefinition
	{
		get
		{
			if (_cacheDefinition == null)
			{
				_cacheDefinition = new ExcelPivotCacheDefinition(base.NameSpaceManager, this, null, 1);
			}
			return _cacheDefinition;
		}
	}

	public ExcelWorksheet WorkSheet { get; set; }

	public ExcelAddressBase Address { get; internal set; }

	public bool DataOnRows
	{
		get
		{
			return GetXmlNodeBool("@dataOnRows");
		}
		set
		{
			SetXmlNodeBool("@dataOnRows", value);
		}
	}

	public bool ApplyNumberFormats
	{
		get
		{
			return GetXmlNodeBool("@applyNumberFormats");
		}
		set
		{
			SetXmlNodeBool("@applyNumberFormats", value);
		}
	}

	public bool ApplyBorderFormats
	{
		get
		{
			return GetXmlNodeBool("@applyBorderFormats");
		}
		set
		{
			SetXmlNodeBool("@applyBorderFormats", value);
		}
	}

	public bool ApplyFontFormats
	{
		get
		{
			return GetXmlNodeBool("@applyFontFormats");
		}
		set
		{
			SetXmlNodeBool("@applyFontFormats", value);
		}
	}

	public bool ApplyPatternFormats
	{
		get
		{
			return GetXmlNodeBool("@applyPatternFormats");
		}
		set
		{
			SetXmlNodeBool("@applyPatternFormats", value);
		}
	}

	public bool ApplyWidthHeightFormats
	{
		get
		{
			return GetXmlNodeBool("@applyWidthHeightFormats");
		}
		set
		{
			SetXmlNodeBool("@applyWidthHeightFormats", value);
		}
	}

	public bool ShowMemberPropertyTips
	{
		get
		{
			return GetXmlNodeBool("@showMemberPropertyTips");
		}
		set
		{
			SetXmlNodeBool("@showMemberPropertyTips", value);
		}
	}

	public bool ShowCalcMember
	{
		get
		{
			return GetXmlNodeBool("@showCalcMbrs");
		}
		set
		{
			SetXmlNodeBool("@showCalcMbrs", value);
		}
	}

	public bool EnableDrill
	{
		get
		{
			return GetXmlNodeBool("@enableDrill", blankValue: true);
		}
		set
		{
			SetXmlNodeBool("@enableDrill", value);
		}
	}

	public bool ShowDrill
	{
		get
		{
			return GetXmlNodeBool("@showDrill", blankValue: true);
		}
		set
		{
			SetXmlNodeBool("@showDrill", value);
		}
	}

	public bool ShowDataTips
	{
		get
		{
			return GetXmlNodeBool("@showDataTips", blankValue: true);
		}
		set
		{
			SetXmlNodeBool("@showDataTips", value, removeIf: true);
		}
	}

	public bool FieldPrintTitles
	{
		get
		{
			return GetXmlNodeBool("@fieldPrintTitles");
		}
		set
		{
			SetXmlNodeBool("@fieldPrintTitles", value);
		}
	}

	public bool ItemPrintTitles
	{
		get
		{
			return GetXmlNodeBool("@itemPrintTitles");
		}
		set
		{
			SetXmlNodeBool("@itemPrintTitles", value);
		}
	}

	[Obsolete("Use correctly spelled property 'ColumnGrandTotals'")]
	public bool ColumGrandTotals
	{
		get
		{
			return ColumnGrandTotals;
		}
		set
		{
			ColumnGrandTotals = value;
		}
	}

	public bool ColumnGrandTotals
	{
		get
		{
			return GetXmlNodeBool("@colGrandTotals");
		}
		set
		{
			SetXmlNodeBool("@colGrandTotals", value);
		}
	}

	public bool RowGrandTotals
	{
		get
		{
			return GetXmlNodeBool("@rowGrandTotals");
		}
		set
		{
			SetXmlNodeBool("@rowGrandTotals", value);
		}
	}

	public bool PrintDrill
	{
		get
		{
			return GetXmlNodeBool("@printDrill");
		}
		set
		{
			SetXmlNodeBool("@printDrill", value);
		}
	}

	public bool ShowError
	{
		get
		{
			return GetXmlNodeBool("@showError");
		}
		set
		{
			SetXmlNodeBool("@showError", value);
		}
	}

	public string ErrorCaption
	{
		get
		{
			return GetXmlNodeString("@errorCaption");
		}
		set
		{
			SetXmlNodeString("@errorCaption", value);
		}
	}

	public string DataCaption
	{
		get
		{
			return GetXmlNodeString("@dataCaption");
		}
		set
		{
			SetXmlNodeString("@dataCaption", value);
		}
	}

	public bool ShowHeaders
	{
		get
		{
			return GetXmlNodeBool("@showHeaders");
		}
		set
		{
			SetXmlNodeBool("@showHeaders", value);
		}
	}

	public int PageWrap
	{
		get
		{
			return GetXmlNodeInt("@pageWrap");
		}
		set
		{
			if (value < 0)
			{
				throw new Exception("Value can't be negative");
			}
			SetXmlNodeString("@pageWrap", value.ToString());
		}
	}

	public bool UseAutoFormatting
	{
		get
		{
			return GetXmlNodeBool("@useAutoFormatting");
		}
		set
		{
			SetXmlNodeBool("@useAutoFormatting", value);
		}
	}

	public bool GridDropZones
	{
		get
		{
			return GetXmlNodeBool("@gridDropZones");
		}
		set
		{
			SetXmlNodeBool("@gridDropZones", value);
		}
	}

	public int Indent
	{
		get
		{
			return GetXmlNodeInt("@indent");
		}
		set
		{
			SetXmlNodeString("@indent", value.ToString());
		}
	}

	public bool OutlineData
	{
		get
		{
			return GetXmlNodeBool("@outlineData");
		}
		set
		{
			SetXmlNodeBool("@outlineData", value);
		}
	}

	public bool Outline
	{
		get
		{
			return GetXmlNodeBool("@outline");
		}
		set
		{
			SetXmlNodeBool("@outline", value);
		}
	}

	public bool MultipleFieldFilters
	{
		get
		{
			return GetXmlNodeBool("@multipleFieldFilters");
		}
		set
		{
			SetXmlNodeBool("@multipleFieldFilters", value);
		}
	}

	public bool Compact
	{
		get
		{
			return GetXmlNodeBool("@compact");
		}
		set
		{
			SetXmlNodeBool("@compact", value);
		}
	}

	public bool CompactData
	{
		get
		{
			return GetXmlNodeBool("@compactData");
		}
		set
		{
			SetXmlNodeBool("@compactData", value);
		}
	}

	public string GrandTotalCaption
	{
		get
		{
			return GetXmlNodeString("@grandTotalCaption");
		}
		set
		{
			SetXmlNodeString("@grandTotalCaption", value);
		}
	}

	public string RowHeaderCaption
	{
		get
		{
			return GetXmlNodeString("@rowHeaderCaption");
		}
		set
		{
			SetXmlNodeString("@rowHeaderCaption", value);
		}
	}

	public string ColumnHeaderCaption
	{
		get
		{
			return GetXmlNodeString("@colHeaderCaption");
		}
		set
		{
			SetXmlNodeString("@colHeaderCaption", value);
		}
	}

	public string MissingCaption
	{
		get
		{
			return GetXmlNodeString("@missingCaption");
		}
		set
		{
			SetXmlNodeString("@missingCaption", value);
		}
	}

	public int FirstHeaderRow
	{
		get
		{
			return GetXmlNodeInt("d:location/@firstHeaderRow");
		}
		set
		{
			SetXmlNodeString("d:location/@firstHeaderRow", value.ToString());
		}
	}

	public int FirstDataRow
	{
		get
		{
			return GetXmlNodeInt("d:location/@firstDataRow");
		}
		set
		{
			SetXmlNodeString("d:location/@firstDataRow", value.ToString());
		}
	}

	public int FirstDataCol
	{
		get
		{
			return GetXmlNodeInt("d:location/@firstDataCol");
		}
		set
		{
			SetXmlNodeString("d:location/@firstDataCol", value.ToString());
		}
	}

	public ExcelPivotTableFieldCollection Fields
	{
		get
		{
			if (_fields == null)
			{
				_fields = new ExcelPivotTableFieldCollection(this, "");
			}
			return _fields;
		}
	}

	public ExcelPivotTableRowColumnFieldCollection RowFields
	{
		get
		{
			if (_rowFields == null)
			{
				_rowFields = new ExcelPivotTableRowColumnFieldCollection(this, "rowFields");
			}
			return _rowFields;
		}
	}

	public ExcelPivotTableRowColumnFieldCollection ColumnFields
	{
		get
		{
			if (_columnFields == null)
			{
				_columnFields = new ExcelPivotTableRowColumnFieldCollection(this, "colFields");
			}
			return _columnFields;
		}
	}

	public ExcelPivotTableDataFieldCollection DataFields
	{
		get
		{
			if (_dataFields == null)
			{
				_dataFields = new ExcelPivotTableDataFieldCollection(this);
			}
			return _dataFields;
		}
	}

	public ExcelPivotTableRowColumnFieldCollection PageFields
	{
		get
		{
			if (_pageFields == null)
			{
				_pageFields = new ExcelPivotTableRowColumnFieldCollection(this, "pageFields");
			}
			return _pageFields;
		}
	}

	public string StyleName
	{
		get
		{
			return GetXmlNodeString("d:pivotTableStyleInfo/@name");
		}
		set
		{
			if (value.StartsWith("PivotStyle"))
			{
				try
				{
					_tableStyle = (TableStyles)Enum.Parse(typeof(TableStyles), value.Substring(10, value.Length - 10), ignoreCase: true);
				}
				catch
				{
					_tableStyle = TableStyles.Custom;
				}
			}
			else if (value == "None")
			{
				_tableStyle = TableStyles.None;
				value = "";
			}
			else
			{
				_tableStyle = TableStyles.Custom;
			}
			SetXmlNodeString("d:pivotTableStyleInfo/@name", value, removeIfBlank: true);
		}
	}

	public TableStyles TableStyle
	{
		get
		{
			return _tableStyle;
		}
		set
		{
			_tableStyle = value;
			if (value != TableStyles.Custom)
			{
				SetXmlNodeString("d:pivotTableStyleInfo/@name", "PivotStyle" + value);
			}
		}
	}

	internal int CacheID
	{
		get
		{
			return GetXmlNodeInt("@cacheId");
		}
		set
		{
			SetXmlNodeString("@cacheId", value.ToString());
		}
	}

	internal ExcelPivotTable(ZipPackageRelationship rel, ExcelWorksheet sheet)
		: base(sheet.NameSpaceManager)
	{
		WorkSheet = sheet;
		PivotTableUri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
		Relationship = rel;
		ZipPackage package = sheet._package.Package;
		Part = package.GetPart(PivotTableUri);
		PivotTableXml = new XmlDocument();
		XmlHelper.LoadXmlSafe(PivotTableXml, Part.GetStream());
		init();
		base.TopNode = PivotTableXml.DocumentElement;
		Address = new ExcelAddressBase(GetXmlNodeString("d:location/@ref"));
		_cacheDefinition = new ExcelPivotCacheDefinition(sheet.NameSpaceManager, this);
		LoadFields();
		foreach (XmlElement item in base.TopNode.SelectNodes("d:rowFields/d:field", base.NameSpaceManager))
		{
			if (int.TryParse(item.GetAttribute("x"), out var result) && result >= 0)
			{
				RowFields.AddInternal(Fields[result]);
			}
			else
			{
				item.ParentNode.RemoveChild(item);
			}
		}
		foreach (XmlElement item2 in base.TopNode.SelectNodes("d:colFields/d:field", base.NameSpaceManager))
		{
			if (int.TryParse(item2.GetAttribute("x"), out var result2) && result2 >= 0)
			{
				ColumnFields.AddInternal(Fields[result2]);
			}
			else
			{
				item2.ParentNode.RemoveChild(item2);
			}
		}
		foreach (XmlElement item3 in base.TopNode.SelectNodes("d:pageFields/d:pageField", base.NameSpaceManager))
		{
			if (int.TryParse(item3.GetAttribute("fld"), out var result3) && result3 >= 0)
			{
				ExcelPivotTableField excelPivotTableField = Fields[result3];
				excelPivotTableField._pageFieldSettings = new ExcelPivotTablePageFieldSettings(base.NameSpaceManager, item3, excelPivotTableField, result3);
				PageFields.AddInternal(excelPivotTableField);
			}
		}
		foreach (XmlElement item4 in base.TopNode.SelectNodes("d:dataFields/d:dataField", base.NameSpaceManager))
		{
			if (int.TryParse(item4.GetAttribute("fld"), out var result4) && result4 >= 0)
			{
				ExcelPivotTableField field = Fields[result4];
				ExcelPivotTableDataField field2 = new ExcelPivotTableDataField(base.NameSpaceManager, item4, field);
				DataFields.AddInternal(field2);
			}
		}
	}

	internal ExcelPivotTable(ExcelWorksheet sheet, ExcelAddressBase address, ExcelRangeBase sourceAddress, string name, int tblId)
		: base(sheet.NameSpaceManager)
	{
		WorkSheet = sheet;
		Address = address;
		ZipPackage package = sheet._package.Package;
		PivotTableXml = new XmlDocument();
		XmlHelper.LoadXmlSafe(PivotTableXml, GetStartXml(name, tblId, address, sourceAddress), Encoding.UTF8);
		base.TopNode = PivotTableXml.DocumentElement;
		PivotTableUri = XmlHelper.GetNewUri(package, "/xl/pivotTables/pivotTable{0}.xml", ref tblId);
		init();
		Part = package.CreatePart(PivotTableUri, "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml");
		PivotTableXml.Save(Part.GetStream());
		Relationship = sheet.Part.CreateRelationship(UriHelper.ResolvePartUri(sheet.WorksheetUri, PivotTableUri), TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable");
		_cacheDefinition = new ExcelPivotCacheDefinition(sheet.NameSpaceManager, this, sourceAddress, tblId);
		_cacheDefinition.Relationship = Part.CreateRelationship(UriHelper.ResolvePartUri(PivotTableUri, _cacheDefinition.CacheDefinitionUri), TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition");
		sheet.Workbook.AddPivotTable(CacheID.ToString(), _cacheDefinition.CacheDefinitionUri);
		LoadFields();
		using ExcelRange excelRange = sheet.Cells[address.Address];
		excelRange.Clear();
	}

	private void init()
	{
		base.SchemaNodeOrder = new string[12]
		{
			"location", "pivotFields", "rowFields", "rowItems", "colFields", "colItems", "pageFields", "pageItems", "dataFields", "dataItems",
			"formats", "pivotTableStyleInfo"
		};
	}

	private void LoadFields()
	{
		int index = 0;
		foreach (XmlElement item in base.TopNode.SelectNodes("d:pivotFields/d:pivotField", base.NameSpaceManager))
		{
			ExcelPivotTableField field = new ExcelPivotTableField(base.NameSpaceManager, item, this, index, index++);
			Fields.AddInternal(field);
		}
		index = 0;
		foreach (XmlElement item2 in _cacheDefinition.TopNode.SelectNodes("d:cacheFields/d:cacheField", base.NameSpaceManager))
		{
			Fields[index++].SetCacheFieldNode(item2);
		}
	}

	private string GetStartXml(string name, int id, ExcelAddressBase address, ExcelAddressBase sourceAddress)
	{
		string text = $"<pivotTableDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" name=\"{ConvertUtil.ExcelEscapeString(name)}\" cacheId=\"{id}\" dataOnRows=\"1\" applyNumberFormats=\"0\" applyBorderFormats=\"0\" applyFontFormats=\"0\" applyPatternFormats=\"0\" applyAlignmentFormats=\"0\" applyWidthHeightFormats=\"1\" dataCaption=\"Data\"  createdVersion=\"4\" showMemberPropertyTips=\"0\" useAutoFormatting=\"1\" itemPrintTitles=\"1\" indent=\"0\" compact=\"0\" compactData=\"0\" gridDropZones=\"1\">";
		text += $"<location ref=\"{address.FirstAddress}\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\" /> ";
		text += $"<pivotFields count=\"{sourceAddress._toCol - sourceAddress._fromCol + 1}\">";
		for (int i = sourceAddress._fromCol; i <= sourceAddress._toCol; i++)
		{
			text += "<pivotField showAll=\"0\" />";
		}
		text += "</pivotFields>";
		text += "<pivotTableStyleInfo name=\"PivotStyleMedium9\" showRowHeaders=\"1\" showColHeaders=\"1\" showRowStripes=\"0\" showColStripes=\"0\" showLastColumn=\"1\" />";
		return text + "</pivotTableDefinition>";
	}

	private string cleanDisplayName(string name)
	{
		return Regex.Replace(name, "[^\\w\\.-_]", "_");
	}
}
