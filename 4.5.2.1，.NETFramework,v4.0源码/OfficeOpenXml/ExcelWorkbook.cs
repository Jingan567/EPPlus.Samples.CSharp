using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;
using OfficeOpenXml.Compatibility;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Packaging.Ionic.Zip;
using OfficeOpenXml.Packaging.Ionic.Zlib;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils;
using OfficeOpenXml.VBA;

namespace OfficeOpenXml;

public sealed class ExcelWorkbook : XmlHelper, IDisposable
{
	internal class SharedStringItem
	{
		internal int pos;

		internal string Text;

		internal bool isRichText;
	}

	internal ExcelPackage _package;

	internal ExcelWorksheets _worksheets;

	private OfficeProperties _properties;

	private ExcelStyles _styles;

	internal Dictionary<string, SharedStringItem> _sharedStrings = new Dictionary<string, SharedStringItem>();

	internal List<SharedStringItem> _sharedStringsList = new List<SharedStringItem>();

	internal ExcelNamedRangeCollection _names;

	internal int _nextDrawingID;

	internal int _nextTableID = int.MinValue;

	internal int _nextPivotTableID = int.MinValue;

	internal XmlNamespaceManager _namespaceManager;

	internal FormulaParser _formulaParser;

	internal FormulaParserManager _parserManager;

	internal CellStore<List<Token>> _formulaTokens;

	private decimal _standardFontWidth = decimal.MinValue;

	private string _fontID = "";

	private ExcelProtection _protection;

	private ExcelWorkbookView _view;

	private ExcelVbaProject _vba;

	private XmlDocument _workbookXml;

	private const string codeModuleNamePath = "d:workbookPr/@codeName";

	private const string date1904Path = "d:workbookPr/@date1904";

	internal const double date1904Offset = 1462.0;

	private bool? date1904Cache;

	private XmlDocument _stylesXml;

	private string CALC_MODE_PATH = "d:calcPr/@calcMode";

	private const string FULL_CALC_ON_LOAD_PATH = "d:calcPr/@fullCalcOnLoad";

	internal List<string> _externalReferences = new List<string>();

	public ExcelWorksheets Worksheets
	{
		get
		{
			if (_worksheets == null)
			{
				XmlNode xmlNode = _workbookXml.DocumentElement.SelectSingleNode("d:sheets", _namespaceManager);
				if (xmlNode == null)
				{
					xmlNode = CreateNode("d:sheets");
				}
				_worksheets = new ExcelWorksheets(_package, _namespaceManager, xmlNode);
			}
			return _worksheets;
		}
	}

	public ExcelNamedRangeCollection Names => _names;

	internal FormulaParser FormulaParser
	{
		get
		{
			if (_formulaParser == null)
			{
				_formulaParser = new FormulaParser(new EpplusExcelDataProvider(_package));
			}
			return _formulaParser;
		}
	}

	public FormulaParserManager FormulaParserManager
	{
		get
		{
			if (_parserManager == null)
			{
				_parserManager = new FormulaParserManager(FormulaParser);
			}
			return _parserManager;
		}
	}

	public decimal MaxFontWidth
	{
		get
		{
			int num = Styles.NamedStyles.FindIndexByID("Normal");
			if (num >= 0)
			{
				if (_standardFontWidth == decimal.MinValue || _fontID != Styles.NamedStyles[num].Style.Font.Id)
				{
					ExcelFont font = Styles.NamedStyles[num].Style.Font;
					try
					{
						_standardFontWidth = GetWidthPixels(font.Name, font.Size);
						_fontID = Styles.NamedStyles[num].Style.Font.Id;
					}
					catch
					{
						_standardFontWidth = (int)((double)font.Size * (2.0 / 3.0));
					}
				}
			}
			else
			{
				_standardFontWidth = 7m;
			}
			return _standardFontWidth;
		}
		set
		{
			_standardFontWidth = value;
		}
	}

	public ExcelProtection Protection
	{
		get
		{
			if (_protection == null)
			{
				_protection = new ExcelProtection(base.NameSpaceManager, base.TopNode, this);
				_protection.SchemaNodeOrder = base.SchemaNodeOrder;
			}
			return _protection;
		}
	}

	public ExcelWorkbookView View
	{
		get
		{
			if (_view == null)
			{
				_view = new ExcelWorkbookView(base.NameSpaceManager, base.TopNode, this);
			}
			return _view;
		}
	}

	public ExcelVbaProject VbaProject
	{
		get
		{
			if (_vba == null && _package.Package.PartExists(new Uri("/xl/vbaProject.bin", UriKind.Relative)))
			{
				_vba = new ExcelVbaProject(this);
			}
			return _vba;
		}
	}

	internal Uri WorkbookUri { get; private set; }

	internal Uri StylesUri { get; private set; }

	internal Uri SharedStringsUri { get; private set; }

	internal ZipPackagePart Part => _package.Package.GetPart(WorkbookUri);

	public XmlDocument WorkbookXml
	{
		get
		{
			if (_workbookXml == null)
			{
				CreateWorkbookXml(_namespaceManager);
			}
			return _workbookXml;
		}
	}

	internal string CodeModuleName
	{
		get
		{
			return GetXmlNodeString("d:workbookPr/@codeName");
		}
		set
		{
			SetXmlNodeString("d:workbookPr/@codeName", value);
		}
	}

	public ExcelVBAModule CodeModule
	{
		get
		{
			if (VbaProject != null)
			{
				return VbaProject.Modules[CodeModuleName];
			}
			return null;
		}
	}

	public bool Date1904
	{
		get
		{
			if (!date1904Cache.HasValue)
			{
				date1904Cache = GetXmlNodeBool("d:workbookPr/@date1904", blankValue: false);
			}
			return date1904Cache.Value;
		}
		set
		{
			if (Date1904 != value)
			{
				foreach (ExcelWorksheet worksheet in Worksheets)
				{
					worksheet.UpdateCellsWithDate1904Setting();
				}
			}
			date1904Cache = value;
			SetXmlNodeBool("d:workbookPr/@date1904", value, removeIf: false);
		}
	}

	public XmlDocument StylesXml
	{
		get
		{
			if (_stylesXml == null)
			{
				if (_package.Package.PartExists(StylesUri))
				{
					_stylesXml = _package.GetXmlFromUri(StylesUri);
				}
				else
				{
					ZipPackagePart zipPackagePart = _package.Package.CreatePart(StylesUri, "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", _package.Compression);
					StringBuilder stringBuilder = new StringBuilder("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
					stringBuilder.Append("<numFmts />");
					stringBuilder.Append("<fonts count=\"1\"><font><sz val=\"11\" /><name val=\"Calibri\" /></font></fonts>");
					stringBuilder.Append("<fills><fill><patternFill patternType=\"none\" /></fill><fill><patternFill patternType=\"gray125\" /></fill></fills>");
					stringBuilder.Append("<borders><border><left /><right /><top /><bottom /><diagonal /></border></borders>");
					stringBuilder.Append("<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" /></cellStyleXfs>");
					stringBuilder.Append("<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" xfId=\"0\" /></cellXfs>");
					stringBuilder.Append("<cellStyles><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\" /></cellStyles>");
					stringBuilder.Append("<dxfs count=\"0\" />");
					stringBuilder.Append("</styleSheet>");
					_stylesXml = new XmlDocument();
					_stylesXml.LoadXml(stringBuilder.ToString());
					StreamWriter writer = new StreamWriter(zipPackagePart.GetStream(FileMode.Create, FileAccess.Write));
					_stylesXml.Save(writer);
					_package.Package.Flush();
					_package.Workbook.Part.CreateRelationship(UriHelper.GetRelativeUri(WorkbookUri, StylesUri), TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
					_package.Package.Flush();
				}
			}
			return _stylesXml;
		}
		set
		{
			_stylesXml = value;
		}
	}

	public ExcelStyles Styles
	{
		get
		{
			if (_styles == null)
			{
				_styles = new ExcelStyles(base.NameSpaceManager, StylesXml, this);
			}
			return _styles;
		}
	}

	public OfficeProperties Properties
	{
		get
		{
			if (_properties == null)
			{
				_properties = new OfficeProperties(_package, base.NameSpaceManager);
			}
			return _properties;
		}
	}

	public ExcelCalcMode CalcMode
	{
		get
		{
			string xmlNodeString = GetXmlNodeString(CALC_MODE_PATH);
			if (!(xmlNodeString == "autoNoTable"))
			{
				if (xmlNodeString == "manual")
				{
					return ExcelCalcMode.Manual;
				}
				return ExcelCalcMode.Automatic;
			}
			return ExcelCalcMode.AutomaticNoTable;
		}
		set
		{
			switch (value)
			{
			case ExcelCalcMode.AutomaticNoTable:
				SetXmlNodeString(CALC_MODE_PATH, "autoNoTable");
				break;
			case ExcelCalcMode.Manual:
				SetXmlNodeString(CALC_MODE_PATH, "manual");
				break;
			default:
				SetXmlNodeString(CALC_MODE_PATH, "auto");
				break;
			}
		}
	}

	public bool FullCalcOnLoad
	{
		get
		{
			return GetXmlNodeBool("d:calcPr/@fullCalcOnLoad");
		}
		set
		{
			SetXmlNodeBool("d:calcPr/@fullCalcOnLoad", value);
		}
	}

	internal ExcelWorkbook(ExcelPackage package, XmlNamespaceManager namespaceManager)
		: base(namespaceManager)
	{
		_package = package;
		WorkbookUri = new Uri("/xl/workbook.xml", UriKind.Relative);
		SharedStringsUri = new Uri("/xl/sharedStrings.xml", UriKind.Relative);
		StylesUri = new Uri("/xl/styles.xml", UriKind.Relative);
		_names = new ExcelNamedRangeCollection(this);
		_namespaceManager = namespaceManager;
		base.TopNode = WorkbookXml.DocumentElement;
		base.SchemaNodeOrder = new string[20]
		{
			"fileVersion", "fileSharing", "workbookPr", "workbookProtection", "bookViews", "sheets", "functionGroups", "functionPrototypes", "externalReferences", "definedNames",
			"calcPr", "oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"
		};
		FullCalcOnLoad = true;
		GetSharedStrings();
	}

	private void GetSharedStrings()
	{
		if (!_package.Package.PartExists(SharedStringsUri))
		{
			return;
		}
		XmlNodeList xmlNodeList = _package.GetXmlFromUri(SharedStringsUri).SelectNodes("//d:sst/d:si", base.NameSpaceManager);
		_sharedStringsList = new List<SharedStringItem>();
		if (xmlNodeList != null)
		{
			foreach (XmlNode item in xmlNodeList)
			{
				XmlNode xmlNode2 = item.SelectSingleNode("d:t", base.NameSpaceManager);
				if (xmlNode2 != null)
				{
					_sharedStringsList.Add(new SharedStringItem
					{
						Text = ConvertUtil.ExcelDecodeString(xmlNode2.InnerText)
					});
				}
				else
				{
					_sharedStringsList.Add(new SharedStringItem
					{
						Text = item.InnerXml,
						isRichText = true
					});
				}
			}
		}
		foreach (ZipPackageRelationship relationship in Part.GetRelationships())
		{
			if (relationship.TargetUri.OriginalString.EndsWith("sharedstrings.xml", StringComparison.OrdinalIgnoreCase))
			{
				Part.DeleteRelationship(relationship.Id);
				break;
			}
		}
		_package.Package.DeletePart(SharedStringsUri);
	}

	internal void GetDefinedNames()
	{
		XmlNodeList xmlNodeList = WorkbookXml.SelectNodes("//d:definedNames/d:definedName", base.NameSpaceManager);
		if (xmlNodeList == null)
		{
			return;
		}
		foreach (XmlElement item in xmlNodeList)
		{
			string text = item.InnerText;
			ExcelWorksheet excelWorksheet;
			if (!int.TryParse(item.GetAttribute("localSheetId"), NumberStyles.Number, CultureInfo.InvariantCulture, out var result))
			{
				result = -1;
				excelWorksheet = null;
			}
			else
			{
				excelWorksheet = Worksheets[result + _package._worksheetAdd];
			}
			ExcelAddressBase.AddressType addressType = ExcelAddressBase.IsValid(text);
			if (text.IndexOf("[") == 0)
			{
				int num = text.IndexOf("[");
				int num2 = text.IndexOf("]", num);
				if (num >= 0 && num2 >= 0 && int.TryParse(text.Substring(num + 1, num2 - num - 1), NumberStyles.Any, CultureInfo.InvariantCulture, out var result2) && result2 > 0 && result2 <= _externalReferences.Count)
				{
					text = text.Substring(0, num) + "[" + _externalReferences[result2 - 1] + "]" + text.Substring(num2 + 1);
				}
			}
			ExcelNamedRange excelNamedRange;
			if (addressType == ExcelAddressBase.AddressType.Invalid || addressType == ExcelAddressBase.AddressType.InternalName || addressType == ExcelAddressBase.AddressType.ExternalName || addressType == ExcelAddressBase.AddressType.Formula || addressType == ExcelAddressBase.AddressType.ExternalAddress)
			{
				ExcelRangeBase range = new ExcelRangeBase(this, excelWorksheet, item.GetAttribute("name"), isName: true);
				excelNamedRange = ((excelWorksheet != null) ? excelWorksheet.Names.Add(item.GetAttribute("name"), range) : _names.Add(item.GetAttribute("name"), range));
				double result3;
				if (ConvertUtil._invariantCompareInfo.IsPrefix(text, "\""))
				{
					excelNamedRange.NameValue = text.Substring(1, text.Length - 2);
				}
				else if (double.TryParse(text, NumberStyles.Number, CultureInfo.InvariantCulture, out result3))
				{
					excelNamedRange.NameValue = result3;
				}
				else
				{
					excelNamedRange.NameFormula = text;
				}
			}
			else
			{
				ExcelAddress excelAddress = new ExcelAddress(text, _package, null);
				if (result > -1)
				{
					excelNamedRange = ((!string.IsNullOrEmpty(excelAddress._ws)) ? Worksheets[result + _package._worksheetAdd].Names.Add(item.GetAttribute("name"), new ExcelRangeBase(this, Worksheets[excelAddress._ws], text, isName: false)) : Worksheets[result + _package._worksheetAdd].Names.Add(item.GetAttribute("name"), new ExcelRangeBase(this, Worksheets[result + _package._worksheetAdd], text, isName: false)));
				}
				else
				{
					ExcelWorksheet xlWorksheet = Worksheets[excelAddress._ws];
					excelNamedRange = _names.Add(item.GetAttribute("name"), new ExcelRangeBase(this, xlWorksheet, text, isName: false));
				}
			}
			if (item.GetAttribute("hidden") == "1" && excelNamedRange != null)
			{
				excelNamedRange.IsNameHidden = true;
			}
			if (!string.IsNullOrEmpty(item.GetAttribute("comment")))
			{
				excelNamedRange.NameComment = item.GetAttribute("comment");
			}
		}
	}

	internal static decimal GetWidthPixels(string fontName, float fontSize)
	{
		Dictionary<float, FontSizeInfo> dictionary = ((!FontSize.FontHeights.ContainsKey(fontName)) ? FontSize.FontHeights["Calibri"] : FontSize.FontHeights[fontName]);
		if (dictionary.ContainsKey(fontSize))
		{
			return Convert.ToDecimal(dictionary[fontSize].Width);
		}
		float num = -1f;
		float num2 = 500f;
		foreach (KeyValuePair<float, FontSizeInfo> item in dictionary)
		{
			if (num < item.Key && item.Key < fontSize)
			{
				num = item.Key;
			}
			if (num2 > item.Key && item.Key > fontSize)
			{
				num2 = item.Key;
			}
		}
		if (num == num2)
		{
			return Convert.ToDecimal(dictionary[num].Width);
		}
		return Convert.ToDecimal(dictionary[num].Height + (dictionary[num2].Height - dictionary[num].Height) * ((fontSize - num) / (num2 - num)));
	}

	public void CreateVBAProject()
	{
		if (_vba != null || _package.Package.PartExists(new Uri("/xl/vbaProject.bin", UriKind.Relative)))
		{
			throw new InvalidOperationException("VBA project already exists.");
		}
		_vba = new ExcelVbaProject(this);
		_vba.Create();
	}

	internal void CodeNameChange(string value)
	{
		CodeModuleName = value;
	}

	internal bool ExistsPivotCache(int cacheID, ref int newID)
	{
		newID = cacheID;
		bool flag = true;
		foreach (ExcelWorksheet worksheet in Worksheets)
		{
			foreach (ExcelPivotTable pivotTable in worksheet.PivotTables)
			{
				if (pivotTable.CacheID == cacheID)
				{
					flag = false;
				}
				if (pivotTable.CacheID >= newID)
				{
					newID = pivotTable.CacheID + 1;
				}
			}
		}
		if (flag)
		{
			newID = cacheID;
		}
		return flag;
	}

	private void CreateWorkbookXml(XmlNamespaceManager namespaceManager)
	{
		if (_package.Package.PartExists(WorkbookUri))
		{
			_workbookXml = _package.GetXmlFromUri(WorkbookUri);
			return;
		}
		ZipPackagePart zipPackagePart = _package.Package.CreatePart(WorkbookUri, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", _package.Compression);
		_workbookXml = new XmlDocument(namespaceManager.NameTable);
		_workbookXml.PreserveWhitespace = false;
		XmlElement xmlElement = _workbookXml.CreateElement("workbook", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
		xmlElement.SetAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
		_workbookXml.AppendChild(xmlElement);
		XmlElement xmlElement2 = _workbookXml.CreateElement("bookViews", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
		xmlElement.AppendChild(xmlElement2);
		XmlElement newChild = _workbookXml.CreateElement("workbookView", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
		xmlElement2.AppendChild(newChild);
		StreamWriter writer = new StreamWriter(zipPackagePart.GetStream(FileMode.Create, FileAccess.Write));
		_workbookXml.Save(writer);
		_package.Package.Flush();
	}

	internal void Save()
	{
		if (Worksheets.Count == 0)
		{
			throw new InvalidOperationException("The workbook must contain at least one worksheet");
		}
		DeleteCalcChain();
		if (_vba == null && !_package.Package.PartExists(new Uri("/xl/vbaProject.bin", UriKind.Relative)))
		{
			if (Part.ContentType != "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
			{
				Part.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
			}
		}
		else if (Part.ContentType != "application/vnd.ms-excel.sheet.macroEnabled.main+xml")
		{
			Part.ContentType = "application/vnd.ms-excel.sheet.macroEnabled.main+xml";
		}
		UpdateDefinedNamesXml();
		if (_workbookXml != null)
		{
			_package.SavePart(WorkbookUri, _workbookXml);
		}
		if (_properties != null)
		{
			_properties.Save();
		}
		Styles.UpdateXml();
		_package.SavePart(StylesUri, _stylesXml);
		bool flag = Protection.LockWindows || Protection.LockStructure;
		foreach (ExcelWorksheet worksheet in Worksheets)
		{
			if (flag && Protection.LockWindows)
			{
				worksheet.View.WindowProtection = true;
			}
			worksheet.Save();
			worksheet.Part.SaveHandler = worksheet.SaveHandler;
		}
		ZipPackagePart zipPackagePart;
		if (_package.Package.PartExists(SharedStringsUri))
		{
			zipPackagePart = _package.Package.GetPart(SharedStringsUri);
		}
		else
		{
			zipPackagePart = _package.Package.CreatePart(SharedStringsUri, "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml", _package.Compression);
			Part.CreateRelationship(UriHelper.GetRelativeUri(WorkbookUri, SharedStringsUri), TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings");
		}
		zipPackagePart.SaveHandler = SaveSharedStringHandler;
		ValidateDataValidations();
		if (_vba != null)
		{
			VbaProject.Save();
		}
	}

	private void DeleteCalcChain()
	{
		Uri uri = new Uri("/xl/calcChain.xml", UriKind.Relative);
		if (!_package.Package.PartExists(uri))
		{
			return;
		}
		Uri uri2 = new Uri("calcChain.xml", UriKind.Relative);
		foreach (ZipPackageRelationship relationship in _package.Workbook.Part.GetRelationships())
		{
			if (relationship.TargetUri == uri2)
			{
				_package.Workbook.Part.DeleteRelationship(relationship.Id);
				break;
			}
		}
		_package.Package.DeletePart(uri);
	}

	private void ValidateDataValidations()
	{
		foreach (ExcelWorksheet worksheet in _package.Workbook.Worksheets)
		{
			if (!(worksheet is ExcelChartsheet))
			{
				worksheet.DataValidations.ValidateAll();
			}
		}
	}

	private void SaveSharedStringHandler(ZipOutputStream stream, CompressionLevel compressionLevel, string fileName)
	{
		stream.CompressionLevel = (OfficeOpenXml.Packaging.Ionic.Zlib.CompressionLevel)compressionLevel;
		stream.PutNextEntry(fileName);
		StringBuilder stringBuilder = new StringBuilder();
		StreamWriter streamWriter = new StreamWriter(stream);
		stringBuilder.AppendFormat("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"{0}\" uniqueCount=\"{0}\">", _sharedStrings.Count);
		foreach (string key in _sharedStrings.Keys)
		{
			if (_sharedStrings[key].isRichText)
			{
				stringBuilder.Append("<si>");
				ConvertUtil.ExcelEncodeString(stringBuilder, key);
				stringBuilder.Append("</si>");
			}
			else
			{
				if (key.Length > 0 && (key[0] == ' ' || key[key.Length - 1] == ' ' || key.Contains("  ") || key.Contains("\t") || key.Contains("\n") || key.Contains("\n")))
				{
					stringBuilder.Append("<si><t xml:space=\"preserve\">");
				}
				else
				{
					stringBuilder.Append("<si><t>");
				}
				ConvertUtil.ExcelEncodeString(stringBuilder, ConvertUtil.ExcelEscapeString(key));
				stringBuilder.Append("</t></si>");
			}
			if (stringBuilder.Length > 6291456)
			{
				streamWriter.Write(stringBuilder.ToString());
				stringBuilder = new StringBuilder();
			}
		}
		stringBuilder.Append("</sst>");
		streamWriter.Write(stringBuilder.ToString());
		streamWriter.Flush();
	}

	private void UpdateDefinedNamesXml()
	{
		try
		{
			XmlNode xmlNode = WorkbookXml.SelectSingleNode("//d:definedNames", base.NameSpaceManager);
			if (!ExistsNames())
			{
				if (xmlNode != null)
				{
					base.TopNode.RemoveChild(xmlNode);
				}
				return;
			}
			if (xmlNode == null)
			{
				CreateNode("d:definedNames");
				xmlNode = WorkbookXml.SelectSingleNode("//d:definedNames", base.NameSpaceManager);
			}
			else
			{
				xmlNode.RemoveAll();
			}
			foreach (ExcelNamedRange name in _names)
			{
				XmlElement xmlElement = WorkbookXml.CreateElement("definedName", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
				xmlNode.AppendChild(xmlElement);
				xmlElement.SetAttribute("name", name.Name);
				if (name.IsNameHidden)
				{
					xmlElement.SetAttribute("hidden", "1");
				}
				if (!string.IsNullOrEmpty(name.NameComment))
				{
					xmlElement.SetAttribute("comment", name.NameComment);
				}
				SetNameElement(name, xmlElement);
			}
			foreach (ExcelWorksheet worksheet in _worksheets)
			{
				if (worksheet is ExcelChartsheet)
				{
					continue;
				}
				foreach (ExcelNamedRange name2 in worksheet.Names)
				{
					XmlElement xmlElement2 = WorkbookXml.CreateElement("definedName", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
					xmlNode.AppendChild(xmlElement2);
					xmlElement2.SetAttribute("name", name2.Name);
					xmlElement2.SetAttribute("localSheetId", name2.LocalSheetId.ToString());
					if (name2.IsNameHidden)
					{
						xmlElement2.SetAttribute("hidden", "1");
					}
					if (!string.IsNullOrEmpty(name2.NameComment))
					{
						xmlElement2.SetAttribute("comment", name2.NameComment);
					}
					SetNameElement(name2, xmlElement2);
				}
			}
		}
		catch (Exception innerException)
		{
			throw new Exception("Internal error updating named ranges ", innerException);
		}
	}

	private void SetNameElement(ExcelNamedRange name, XmlElement elem)
	{
		if (name.IsName)
		{
			if (string.IsNullOrEmpty(name.NameFormula))
			{
				if (TypeCompat.IsPrimitive(name.NameValue) || name.NameValue is double || name.NameValue is decimal)
				{
					elem.InnerText = Convert.ToDouble(name.NameValue, CultureInfo.InvariantCulture).ToString("R15", CultureInfo.InvariantCulture);
				}
				else if (name.NameValue is DateTime)
				{
					elem.InnerText = ((DateTime)name.NameValue).ToOADate().ToString(CultureInfo.InvariantCulture);
				}
				else
				{
					elem.InnerText = "\"" + name.NameValue.ToString() + "\"";
				}
			}
			else
			{
				elem.InnerText = name.NameFormula;
			}
		}
		else
		{
			elem.InnerText = name.FullAddressAbsolute;
		}
	}

	private bool ExistsNames()
	{
		if (_names.Count == 0)
		{
			foreach (ExcelWorksheet worksheet in Worksheets)
			{
				if (!(worksheet is ExcelChartsheet) && worksheet.Names.Count > 0)
				{
					return true;
				}
			}
			return false;
		}
		return true;
	}

	internal bool ExistsTableName(string Name)
	{
		foreach (ExcelWorksheet worksheet in Worksheets)
		{
			if (worksheet.Tables._tableNames.ContainsKey(Name))
			{
				return true;
			}
		}
		return false;
	}

	internal bool ExistsPivotTableName(string Name)
	{
		foreach (ExcelWorksheet worksheet in Worksheets)
		{
			if (worksheet.PivotTables._pivotTableNames.ContainsKey(Name))
			{
				return true;
			}
		}
		return false;
	}

	internal void AddPivotTable(string cacheID, Uri defUri)
	{
		CreateNode("d:pivotCaches");
		XmlElement xmlElement = WorkbookXml.CreateElement("pivotCache", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
		xmlElement.SetAttribute("cacheId", cacheID);
		ZipPackageRelationship zipPackageRelationship = Part.CreateRelationship(UriHelper.ResolvePartUri(WorkbookUri, defUri), TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition");
		xmlElement.SetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", zipPackageRelationship.Id);
		WorkbookXml.SelectSingleNode("//d:pivotCaches", base.NameSpaceManager).AppendChild(xmlElement);
	}

	internal void GetExternalReferences()
	{
		XmlNodeList xmlNodeList = WorkbookXml.SelectNodes("//d:externalReferences/d:externalReference", base.NameSpaceManager);
		if (xmlNodeList == null)
		{
			return;
		}
		foreach (XmlElement item in xmlNodeList)
		{
			string attribute = item.GetAttribute("r:id");
			ZipPackageRelationship relationship = Part.GetRelationship(attribute);
			ZipPackagePart part = _package.Package.GetPart(UriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri));
			XmlDocument xmlDocument = new XmlDocument();
			XmlHelper.LoadXmlSafe(xmlDocument, part.GetStream());
			if (xmlDocument.SelectSingleNode("//d:externalBook", base.NameSpaceManager) is XmlElement xmlElement)
			{
				string attribute2 = xmlElement.GetAttribute("r:id");
				ZipPackageRelationship relationship2 = part.GetRelationship(attribute2);
				if (relationship2 != null)
				{
					_externalReferences.Add(relationship2.TargetUri.OriginalString);
				}
			}
		}
	}

	public void Dispose()
	{
		if (_sharedStrings != null)
		{
			_sharedStrings.Clear();
			_sharedStrings = null;
		}
		if (_sharedStringsList != null)
		{
			_sharedStringsList.Clear();
			_sharedStringsList = null;
		}
		_vba = null;
		if (_worksheets != null)
		{
			_worksheets.Dispose();
			_worksheets = null;
		}
		_package = null;
		_properties = null;
		if (_formulaParser != null)
		{
			_formulaParser.Dispose();
			_formulaParser = null;
		}
	}

	internal void ReadAllTables()
	{
		if (_nextTableID > 0)
		{
			return;
		}
		_nextTableID = 1;
		_nextPivotTableID = 1;
		foreach (ExcelWorksheet worksheet in Worksheets)
		{
			if (worksheet is ExcelChartsheet)
			{
				continue;
			}
			foreach (ExcelTable table in worksheet.Tables)
			{
				if (table.Id >= _nextTableID)
				{
					_nextTableID = table.Id + 1;
				}
			}
			foreach (ExcelPivotTable pivotTable in worksheet.PivotTables)
			{
				if (pivotTable.CacheID >= _nextPivotTableID)
				{
					_nextPivotTableID = pivotTable.CacheID + 1;
				}
			}
		}
	}
}
