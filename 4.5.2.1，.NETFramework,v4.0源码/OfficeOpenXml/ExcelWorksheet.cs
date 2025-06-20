using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using OfficeOpenXml.Compatibility;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.DataValidation.Contracts;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Packaging.Ionic.Zip;
using OfficeOpenXml.Packaging.Ionic.Zlib;
using OfficeOpenXml.Sparkline;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils;
using OfficeOpenXml.VBA;

namespace OfficeOpenXml;

public class ExcelWorksheet : XmlHelper, IEqualityComparer<ExcelWorksheet>, IDisposable
{
	internal class Formulas
	{
		private ISourceCodeTokenizer _tokenizer;

		internal int Index { get; set; }

		internal string Address { get; set; }

		internal bool IsArray { get; set; }

		public string Formula { get; set; }

		public int StartRow { get; set; }

		public int StartCol { get; set; }

		private IEnumerable<Token> Tokens { get; set; }

		public Formulas(ISourceCodeTokenizer tokenizer)
		{
			_tokenizer = tokenizer;
		}

		internal string GetFormula(int row, int column, string worksheet)
		{
			if (StartRow == row && StartCol == column)
			{
				return Formula;
			}
			if (Tokens == null)
			{
				Tokens = _tokenizer.Tokenize(Formula, worksheet);
			}
			string text = "";
			foreach (Token token in Tokens)
			{
				if (token.TokenType == TokenType.ExcelAddress)
				{
					ExcelFormulaAddress excelFormulaAddress = new ExcelFormulaAddress(token.Value);
					text += ((!string.IsNullOrEmpty(excelFormulaAddress._wb) || !string.IsNullOrEmpty(excelFormulaAddress._ws)) ? token.Value : excelFormulaAddress.GetOffset(row - StartRow, column - StartCol));
				}
				else
				{
					text += token.Value;
				}
			}
			return text;
		}

		public Formulas Clone()
		{
			return new Formulas(_tokenizer)
			{
				Index = Index,
				Address = Address,
				IsArray = IsArray,
				Formula = Formula,
				StartRow = StartRow,
				StartCol = StartCol
			};
		}
	}

	public class MergeCellsCollection : IEnumerable<string>, IEnumerable
	{
		internal CellStore<int> _cells = new CellStore<int>();

		private List<string> _list = new List<string>();

		internal List<string> List => _list;

		public string this[int row, int column]
		{
			get
			{
				int value = -1;
				if (_cells.Exists(row, column, ref value) && value >= 0 && value < List.Count)
				{
					return List[value];
				}
				return null;
			}
		}

		public string this[int index] => _list[index];

		public int Count => _list.Count;

		internal MergeCellsCollection()
		{
		}

		internal void Add(ExcelAddressBase address, bool doValidate)
		{
			int num = 0;
			if (doValidate && !Validate(address))
			{
				throw new ArgumentException("Can't merge and already merged range");
			}
			lock (this)
			{
				num = _list.Count;
				_list.Add(address.Address);
				SetIndex(address, num);
			}
		}

		private bool Validate(ExcelAddressBase address)
		{
			int value = 0;
			if (_cells.Exists(address._fromRow, address._fromCol, ref value))
			{
				if (value >= 0 && value < _list.Count && _list[value] != null && address.Address == _list[value])
				{
					return true;
				}
				return false;
			}
			CellsStoreEnumerator<int> cellsStoreEnumerator = new CellsStoreEnumerator<int>(_cells, address._fromRow, address._fromCol, address._toRow, address._toCol);
			if (cellsStoreEnumerator.Next())
			{
				return false;
			}
			cellsStoreEnumerator = new CellsStoreEnumerator<int>(_cells, 0, address._fromCol, 0, address._toCol);
			if (cellsStoreEnumerator.Next())
			{
				return false;
			}
			cellsStoreEnumerator = new CellsStoreEnumerator<int>(_cells, address._fromRow, 0, address._toRow, 0);
			if (cellsStoreEnumerator.Next())
			{
				return false;
			}
			return true;
		}

		internal void SetIndex(ExcelAddressBase address, int ix)
		{
			if (address._fromRow == 1 && address._toRow == 1048576)
			{
				for (int i = address._fromCol; i <= address._toCol; i++)
				{
					_cells.SetValue(0, i, ix);
				}
				return;
			}
			if (address._fromCol == 1 && address._toCol == 16384)
			{
				for (int j = address._fromRow; j <= address._toRow; j++)
				{
					_cells.SetValue(j, 0, ix);
				}
				return;
			}
			for (int k = address._fromCol; k <= address._toCol; k++)
			{
				for (int l = address._fromRow; l <= address._toRow; l++)
				{
					_cells.SetValue(l, k, ix);
				}
			}
		}

		internal void Remove(string Item)
		{
			_list.Remove(Item);
		}

		public IEnumerator<string> GetEnumerator()
		{
			return _list.GetEnumerator();
		}

		IEnumerator IEnumerable.GetEnumerator()
		{
			return _list.GetEnumerator();
		}

		internal void Clear(ExcelAddressBase Destination)
		{
			CellsStoreEnumerator<int> cellsStoreEnumerator = new CellsStoreEnumerator<int>(_cells, Destination._fromRow, Destination._fromCol, Destination._toRow, Destination._toCol);
			HashSet<int> hashSet = new HashSet<int>();
			while (cellsStoreEnumerator.Next())
			{
				int value = cellsStoreEnumerator.Value;
				if (!hashSet.Contains(value) && _list[value] != null)
				{
					ExcelAddressBase excelAddressBase = new ExcelAddressBase(_list[value]);
					if (Destination.Collide(excelAddressBase) != ExcelAddressBase.eAddressCollition.Inside && Destination.Collide(excelAddressBase) != ExcelAddressBase.eAddressCollition.Equal)
					{
						throw new InvalidOperationException($"Can't delete/overwrite merged cells. A range is partly merged with the another merged range. {excelAddressBase._address}");
					}
					hashSet.Add(value);
				}
			}
			_cells.Clear(Destination._fromRow, Destination._fromCol, Destination._toRow - Destination._fromRow + 1, Destination._toCol - Destination._fromCol + 1);
			foreach (int item in hashSet)
			{
				_list[item] = null;
			}
		}
	}

	internal CellStore<ExcelCoreValue> _values;

	internal CellStore<object> _formulas;

	internal FlagCellStore _flags;

	internal CellStore<List<Token>> _formulaTokens;

	internal CellStore<Uri> _hyperLinks;

	internal CellStore<int> _commentsStore;

	internal Dictionary<int, Formulas> _sharedFormulas = new Dictionary<int, Formulas>();

	internal int _minCol = 16384;

	internal int _maxCol;

	internal ExcelPackage _package;

	private Uri _worksheetUri;

	private string _name;

	private int _sheetID;

	private int _positionID;

	private string _relationshipID;

	private XmlDocument _worksheetXml;

	internal ExcelWorksheetView _sheetView;

	internal ExcelHeaderFooter _headerFooter;

	internal ExcelNamedRangeCollection _names;

	private double _defaultRowHeight = double.NaN;

	private const string outLineSummaryBelowPath = "d:sheetPr/d:outlinePr/@summaryBelow";

	private const string outLineSummaryRightPath = "d:sheetPr/d:outlinePr/@summaryRight";

	private const string outLineApplyStylePath = "d:sheetPr/d:outlinePr/@applyStyles";

	private const string tabColorPath = "d:sheetPr/d:tabColor/@rgb";

	private const string codeModuleNamePath = "d:sheetPr/@codeName";

	internal ExcelVmlDrawingCommentCollection _vmlDrawings;

	internal ExcelCommentCollection _comments;

	private const int BLOCKSIZE = 8192;

	private MergeCellsCollection _mergedCells = new MergeCellsCollection();

	private Dictionary<int, int> columnStyles;

	private ExcelSheetProtection _protection;

	private ExcelProtectedRangeCollection _protectedRanges;

	private ExcelDrawings _drawings;

	private ExcelSparklineGroupCollection _sparklineGroups;

	private ExcelTableCollection _tables;

	internal ExcelPivotTableCollection _pivotTables;

	private ExcelConditionalFormattingCollection _conditionalFormatting;

	private ExcelDataValidationCollection _dataValidation;

	private ExcelBackgroundImage _backgroundImage;

	private static CellStore<ExcelCoreValue>.SetValueDelegate _setValueInnerUpdateDelegate = SetValueInnerUpdate;

	internal Uri WorksheetUri => _worksheetUri;

	internal ZipPackagePart Part => _package.Package.GetPart(WorksheetUri);

	internal string RelationshipID => _relationshipID;

	internal int SheetID => _sheetID;

	internal int PositionID
	{
		get
		{
			return _positionID;
		}
		set
		{
			_positionID = value;
		}
	}

	public int Index => _positionID;

	public ExcelAddressBase AutoFilterAddress
	{
		get
		{
			CheckSheetType();
			string xmlNodeString = GetXmlNodeString("d:autoFilter/@ref");
			if (xmlNodeString == "")
			{
				return null;
			}
			return new ExcelAddressBase(xmlNodeString);
		}
		internal set
		{
			CheckSheetType();
			if (value == null)
			{
				DeleteAllNode("d:autoFilter/@ref");
			}
			else
			{
				SetXmlNodeString("d:autoFilter/@ref", value.Address);
			}
		}
	}

	public ExcelWorksheetView View
	{
		get
		{
			if (_sheetView == null)
			{
				XmlNode xmlNode = base.TopNode.SelectSingleNode("d:sheetViews/d:sheetView", base.NameSpaceManager);
				if (xmlNode == null)
				{
					CreateNode("d:sheetViews/d:sheetView");
					xmlNode = base.TopNode.SelectSingleNode("d:sheetViews/d:sheetView", base.NameSpaceManager);
				}
				_sheetView = new ExcelWorksheetView(base.NameSpaceManager, xmlNode, this);
			}
			return _sheetView;
		}
	}

	public string Name
	{
		get
		{
			return _name;
		}
		set
		{
			if (value == _name)
			{
				return;
			}
			value = _package.Workbook.Worksheets.ValidateFixSheetName(value);
			foreach (ExcelWorksheet worksheet in Workbook.Worksheets)
			{
				if (worksheet.PositionID != PositionID && worksheet.Name.Equals(value, StringComparison.OrdinalIgnoreCase))
				{
					throw new ArgumentException("Worksheet name must be unique");
				}
			}
			_package.Workbook.SetXmlNodeString($"d:sheets/d:sheet[@sheetId={_sheetID}]/@name", value);
			ChangeNames(value);
			_name = value;
		}
	}

	public ExcelNamedRangeCollection Names
	{
		get
		{
			CheckSheetType();
			return _names;
		}
	}

	public eWorkSheetHidden Hidden
	{
		get
		{
			string xmlNodeString = _package.Workbook.GetXmlNodeString($"d:sheets/d:sheet[@sheetId={_sheetID}]/@state");
			if (xmlNodeString == "hidden")
			{
				return eWorkSheetHidden.Hidden;
			}
			if (xmlNodeString == "veryHidden")
			{
				return eWorkSheetHidden.VeryHidden;
			}
			return eWorkSheetHidden.Visible;
		}
		set
		{
			if (value == eWorkSheetHidden.Visible)
			{
				_package.Workbook.DeleteNode($"d:sheets/d:sheet[@sheetId={_sheetID}]/@state");
				return;
			}
			string text = value.ToString();
			text = text.Substring(0, 1).ToLowerInvariant() + text.Substring(1);
			_package.Workbook.SetXmlNodeString($"d:sheets/d:sheet[@sheetId={_sheetID}]/@state", text);
		}
	}

	public double DefaultRowHeight
	{
		get
		{
			CheckSheetType();
			_defaultRowHeight = GetXmlNodeDouble("d:sheetFormatPr/@defaultRowHeight");
			if (double.IsNaN(_defaultRowHeight) || !CustomHeight)
			{
				_defaultRowHeight = GetRowHeightFromNormalStyle();
			}
			return _defaultRowHeight;
		}
		set
		{
			CheckSheetType();
			_defaultRowHeight = value;
			if (double.IsNaN(value))
			{
				DeleteNode("d:sheetFormatPr/@defaultRowHeight");
				return;
			}
			SetXmlNodeString("d:sheetFormatPr/@defaultRowHeight", value.ToString(CultureInfo.InvariantCulture));
			GetRowHeightFromNormalStyle();
			CustomHeight = true;
		}
	}

	public bool CustomHeight
	{
		get
		{
			return GetXmlNodeBool("d:sheetFormatPr/@customHeight");
		}
		set
		{
			SetXmlNodeBool("d:sheetFormatPr/@customHeight", value);
		}
	}

	public double DefaultColWidth
	{
		get
		{
			CheckSheetType();
			double xmlNodeDouble = GetXmlNodeDouble("d:sheetFormatPr/@defaultColWidth");
			if (double.IsNaN(xmlNodeDouble))
			{
				double num = Convert.ToDouble(Workbook.MaxFontWidth);
				double num2 = num * 7.0;
				double num3 = Math.Truncate(num / 4.0 + 0.999) * 2.0 + 1.0;
				if (num3 < 5.0)
				{
					num3 = 5.0;
				}
				for (; Math.Truncate((num2 - num3) / num * 100.0 + 0.5) / 100.0 < 8.0; num2 += 1.0)
				{
				}
				num2 = ((num2 % 8.0 == 0.0) ? num2 : (8.0 - num2 % 8.0 + num2));
				return Math.Truncate((Math.Truncate((num2 - num3) / num * 100.0 + 0.5) / 100.0 * num + num3) / num * 256.0) / 256.0;
			}
			return xmlNodeDouble;
		}
		set
		{
			CheckSheetType();
			SetXmlNodeString("d:sheetFormatPr/@defaultColWidth", value.ToString(CultureInfo.InvariantCulture));
			if (double.IsNaN(GetXmlNodeDouble("d:sheetFormatPr/@defaultRowHeight")))
			{
				SetXmlNodeString("d:sheetFormatPr/@defaultRowHeight", GetRowHeightFromNormalStyle().ToString(CultureInfo.InvariantCulture));
			}
		}
	}

	public bool OutLineSummaryBelow
	{
		get
		{
			CheckSheetType();
			return GetXmlNodeBool("d:sheetPr/d:outlinePr/@summaryBelow");
		}
		set
		{
			CheckSheetType();
			SetXmlNodeString("d:sheetPr/d:outlinePr/@summaryBelow", value ? "1" : "0");
		}
	}

	public bool OutLineSummaryRight
	{
		get
		{
			CheckSheetType();
			return GetXmlNodeBool("d:sheetPr/d:outlinePr/@summaryRight");
		}
		set
		{
			CheckSheetType();
			SetXmlNodeString("d:sheetPr/d:outlinePr/@summaryRight", value ? "1" : "0");
		}
	}

	public bool OutLineApplyStyle
	{
		get
		{
			CheckSheetType();
			return GetXmlNodeBool("d:sheetPr/d:outlinePr/@applyStyles");
		}
		set
		{
			CheckSheetType();
			SetXmlNodeString("d:sheetPr/d:outlinePr/@applyStyles", value ? "1" : "0");
		}
	}

	public Color TabColor
	{
		get
		{
			string xmlNodeString = GetXmlNodeString("d:sheetPr/d:tabColor/@rgb");
			if (xmlNodeString == "")
			{
				return Color.Empty;
			}
			return Color.FromArgb(int.Parse(xmlNodeString, NumberStyles.AllowHexSpecifier));
		}
		set
		{
			SetXmlNodeString("d:sheetPr/d:tabColor/@rgb", value.ToArgb().ToString("X"));
		}
	}

	internal string CodeModuleName
	{
		get
		{
			return GetXmlNodeString("d:sheetPr/@codeName");
		}
		set
		{
			SetXmlNodeString("d:sheetPr/@codeName", value);
		}
	}

	public ExcelVBAModule CodeModule
	{
		get
		{
			if (_package.Workbook.VbaProject != null)
			{
				return _package.Workbook.VbaProject.Modules[CodeModuleName];
			}
			return null;
		}
	}

	public XmlDocument WorksheetXml => _worksheetXml;

	internal ExcelVmlDrawingCommentCollection VmlDrawingsComments
	{
		get
		{
			if (_vmlDrawings == null)
			{
				CreateVmlCollection();
			}
			return _vmlDrawings;
		}
	}

	public ExcelCommentCollection Comments
	{
		get
		{
			CheckSheetType();
			if (_comments == null)
			{
				CreateVmlCollection();
				_comments = new ExcelCommentCollection(_package, this, base.NameSpaceManager);
			}
			return _comments;
		}
	}

	public ExcelHeaderFooter HeaderFooter
	{
		get
		{
			if (_headerFooter == null)
			{
				XmlNode xmlNode = base.TopNode.SelectSingleNode("d:headerFooter", base.NameSpaceManager);
				if (xmlNode == null)
				{
					xmlNode = CreateNode("d:headerFooter");
				}
				_headerFooter = new ExcelHeaderFooter(base.NameSpaceManager, xmlNode, this);
			}
			return _headerFooter;
		}
	}

	public ExcelPrinterSettings PrinterSettings => new ExcelPrinterSettings(base.NameSpaceManager, base.TopNode, this)
	{
		SchemaNodeOrder = base.SchemaNodeOrder
	};

	public ExcelRange Cells
	{
		get
		{
			CheckSheetType();
			return new ExcelRange(this, 1, 1, 1048576, 16384);
		}
	}

	public ExcelRange SelectedRange
	{
		get
		{
			CheckSheetType();
			return new ExcelRange(this, View.SelectedRange);
		}
	}

	public MergeCellsCollection MergedCells
	{
		get
		{
			CheckSheetType();
			return _mergedCells;
		}
	}

	public ExcelAddressBase Dimension
	{
		get
		{
			CheckSheetType();
			if (_values.GetDimension(out var fromRow, out var fromCol, out var toRow, out var toCol))
			{
				return new ExcelAddressBase(fromRow, fromCol, toRow, toCol)
				{
					_ws = Name
				};
			}
			return null;
		}
	}

	public ExcelSheetProtection Protection
	{
		get
		{
			if (_protection == null)
			{
				_protection = new ExcelSheetProtection(base.NameSpaceManager, base.TopNode, this);
			}
			return _protection;
		}
	}

	public ExcelProtectedRangeCollection ProtectedRanges
	{
		get
		{
			if (_protectedRanges == null)
			{
				_protectedRanges = new ExcelProtectedRangeCollection(base.NameSpaceManager, base.TopNode, this);
			}
			return _protectedRanges;
		}
	}

	public ExcelDrawings Drawings
	{
		get
		{
			if (_drawings == null)
			{
				_drawings = new ExcelDrawings(_package, this);
			}
			return _drawings;
		}
	}

	public ExcelSparklineGroupCollection SparklineGroups
	{
		get
		{
			if (_sparklineGroups == null)
			{
				_sparklineGroups = new ExcelSparklineGroupCollection(this);
			}
			return _sparklineGroups;
		}
	}

	public ExcelTableCollection Tables
	{
		get
		{
			CheckSheetType();
			if (Workbook._nextTableID == int.MinValue)
			{
				Workbook.ReadAllTables();
			}
			if (_tables == null)
			{
				_tables = new ExcelTableCollection(this);
			}
			return _tables;
		}
	}

	public ExcelPivotTableCollection PivotTables
	{
		get
		{
			CheckSheetType();
			if (_pivotTables == null)
			{
				if (Workbook._nextPivotTableID == int.MinValue)
				{
					Workbook.ReadAllTables();
				}
				_pivotTables = new ExcelPivotTableCollection(this);
			}
			return _pivotTables;
		}
	}

	public ExcelConditionalFormattingCollection ConditionalFormatting
	{
		get
		{
			CheckSheetType();
			if (_conditionalFormatting == null)
			{
				_conditionalFormatting = new ExcelConditionalFormattingCollection(this);
			}
			return _conditionalFormatting;
		}
	}

	public ExcelDataValidationCollection DataValidations
	{
		get
		{
			CheckSheetType();
			if (_dataValidation == null)
			{
				_dataValidation = new ExcelDataValidationCollection(this);
			}
			return _dataValidation;
		}
	}

	public ExcelBackgroundImage BackgroundImage
	{
		get
		{
			if (_backgroundImage == null)
			{
				_backgroundImage = new ExcelBackgroundImage(base.NameSpaceManager, base.TopNode, this);
			}
			return _backgroundImage;
		}
	}

	public ExcelWorkbook Workbook => _package.Workbook;

	public ExcelWorksheet(XmlNamespaceManager ns, ExcelPackage excelPackage, string relID, Uri uriWorksheet, string sheetName, int sheetID, int positionID, eWorkSheetHidden hide)
		: base(ns, null)
	{
		base.SchemaNodeOrder = new string[42]
		{
			"sheetPr", "tabColor", "outlinePr", "pageSetUpPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetProtection",
			"protectedRanges", "scenarios", "autoFilter", "sortState", "dataConsolidate", "customSheetViews", "customSheetViews", "mergeCells", "phoneticPr", "conditionalFormatting",
			"dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "linePrint", "rowBreaks", "colBreaks", "customProperties",
			"cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "activeXControls", "webPublishItems",
			"tableParts", "extLst"
		};
		_package = excelPackage;
		_relationshipID = relID;
		_worksheetUri = uriWorksheet;
		_name = sheetName;
		_sheetID = sheetID;
		_positionID = positionID;
		Hidden = hide;
		_values = new CellStore<ExcelCoreValue>();
		_formulas = new CellStore<object>();
		_flags = new FlagCellStore();
		_commentsStore = new CellStore<int>();
		_hyperLinks = new CellStore<Uri>();
		_names = new ExcelNamedRangeCollection(Workbook, this);
		CreateXml();
		base.TopNode = _worksheetXml.DocumentElement;
	}

	internal void CheckSheetType()
	{
		if (this is ExcelChartsheet)
		{
			throw new NotSupportedException("This property or method is not supported for a Chartsheet");
		}
	}

	private void ChangeNames(string value)
	{
		foreach (ExcelNamedRange name in Workbook.Names)
		{
			if (string.IsNullOrEmpty(name.NameFormula) && name.NameValue == null)
			{
				name.ChangeWorksheet(_name, value);
			}
		}
		foreach (ExcelWorksheet worksheet in Workbook.Worksheets)
		{
			if (worksheet is ExcelChartsheet)
			{
				continue;
			}
			foreach (ExcelNamedRange name2 in worksheet.Names)
			{
				if (string.IsNullOrEmpty(name2.NameFormula) && name2.NameValue == null)
				{
					name2.ChangeWorksheet(_name, value);
				}
			}
			worksheet.UpdateCrossSheetReferenceNames(_name, value);
		}
	}

	private double GetRowHeightFromNormalStyle()
	{
		int num = Workbook.Styles.NamedStyles.FindIndexByID("Normal");
		if (num >= 0)
		{
			ExcelFont font = Workbook.Styles.NamedStyles[num].Style.Font;
			return (double)ExcelFontXml.GetFontHeight(font.Name, font.Size) * 0.75;
		}
		return 15.0;
	}

	internal void CodeNameChange(string value)
	{
		CodeModuleName = value;
	}

	private void CreateVmlCollection()
	{
		XmlNode xmlNode = _worksheetXml.DocumentElement.SelectSingleNode("d:legacyDrawing/@r:id", base.NameSpaceManager);
		if (xmlNode == null)
		{
			_vmlDrawings = new ExcelVmlDrawingCommentCollection(_package, this, null);
		}
		else if (Part.RelationshipExists(xmlNode.Value))
		{
			ZipPackageRelationship relationship = Part.GetRelationship(xmlNode.Value);
			Uri uri = UriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
			_vmlDrawings = new ExcelVmlDrawingCommentCollection(_package, this, uri);
			_vmlDrawings.RelId = relationship.Id;
		}
	}

	private void CreateXml()
	{
		_worksheetXml = new XmlDocument();
		_worksheetXml.PreserveWhitespace = false;
		ZipPackagePart part = _package.Package.GetPart(WorksheetUri);
		string text = "";
		bool doAdjustDrawings = _package.DoAdjustDrawings;
		_package.DoAdjustDrawings = false;
		Stream stream = part.GetStream();
		XmlTextReader xr = new XmlTextReader(stream)
		{
			ProhibitDtd = true,
			WhitespaceHandling = WhitespaceHandling.None
		};
		LoadColumns(xr);
		long position = stream.Position;
		LoadCells(xr);
		int attributeLength = GetAttributeLength(xr);
		long end = stream.Position - attributeLength;
		LoadMergeCells(xr);
		LoadHyperLinks(xr);
		LoadRowPageBreakes(xr);
		LoadColPageBreakes(xr);
		stream.Seek(0L, SeekOrigin.Begin);
		text = GetWorkSheetXml(stream, position, end, out var encoding);
		stream.Dispose();
		part.Stream = new MemoryStream();
		if (text[0] != '<')
		{
			XmlHelper.LoadXmlSafe(_worksheetXml, text.Substring(1, text.Length - 1), encoding);
		}
		else
		{
			XmlHelper.LoadXmlSafe(_worksheetXml, text, encoding);
		}
		_package.DoAdjustDrawings = doAdjustDrawings;
		ClearNodes();
	}

	private int GetAttributeLength(XmlReader xr)
	{
		if (xr.NodeType != XmlNodeType.Element)
		{
			return 0;
		}
		int num = 0;
		for (int i = 0; i < xr.AttributeCount; i++)
		{
			string attribute = xr.GetAttribute(i);
			num += ((!string.IsNullOrEmpty(attribute)) ? attribute.Length : 0);
		}
		return num;
	}

	private void LoadRowPageBreakes(XmlReader xr)
	{
		if (!ReadUntil(xr, "rowBreaks", "colBreaks"))
		{
			return;
		}
		while (xr.Read() && xr.LocalName == "brk")
		{
			if (xr.NodeType == XmlNodeType.Element && int.TryParse(xr.GetAttribute("id"), NumberStyles.Number, CultureInfo.InvariantCulture, out var result))
			{
				Row(result).PageBreak = true;
			}
		}
	}

	private void LoadColPageBreakes(XmlReader xr)
	{
		if (!ReadUntil(xr, "colBreaks"))
		{
			return;
		}
		while (xr.Read() && xr.LocalName == "brk")
		{
			if (xr.NodeType == XmlNodeType.Element && int.TryParse(xr.GetAttribute("id"), NumberStyles.Number, CultureInfo.InvariantCulture, out var result))
			{
				Column(result).PageBreak = true;
			}
		}
	}

	private void ClearNodes()
	{
		if (_worksheetXml.SelectSingleNode("//d:cols", base.NameSpaceManager) != null)
		{
			_worksheetXml.SelectSingleNode("//d:cols", base.NameSpaceManager).RemoveAll();
		}
		if (_worksheetXml.SelectSingleNode("//d:mergeCells", base.NameSpaceManager) != null)
		{
			_worksheetXml.SelectSingleNode("//d:mergeCells", base.NameSpaceManager).RemoveAll();
		}
		if (_worksheetXml.SelectSingleNode("//d:hyperlinks", base.NameSpaceManager) != null)
		{
			_worksheetXml.SelectSingleNode("//d:hyperlinks", base.NameSpaceManager).RemoveAll();
		}
		if (_worksheetXml.SelectSingleNode("//d:rowBreaks", base.NameSpaceManager) != null)
		{
			_worksheetXml.SelectSingleNode("//d:rowBreaks", base.NameSpaceManager).RemoveAll();
		}
		if (_worksheetXml.SelectSingleNode("//d:colBreaks", base.NameSpaceManager) != null)
		{
			_worksheetXml.SelectSingleNode("//d:colBreaks", base.NameSpaceManager).RemoveAll();
		}
	}

	private string GetWorkSheetXml(Stream stream, long start, long end, out Encoding encoding)
	{
		StreamReader streamReader = new StreamReader(stream);
		int num = 0;
		StringBuilder stringBuilder = new StringBuilder();
		do
		{
			int num2 = (int)((stream.Length < 8192) ? stream.Length : 8192);
			char[] array = new char[num2];
			int charCount = streamReader.ReadBlock(array, 0, num2);
			stringBuilder.Append(array, 0, charCount);
			num += num2;
		}
		while (num < start + 20 && num < end);
		Match match = Regex.Match(stringBuilder.ToString(), string.Format("(<[^>]*{0}[^>]*>)", "sheetData"));
		if (!match.Success)
		{
			encoding = streamReader.CurrentEncoding;
			return stringBuilder.ToString();
		}
		string text = stringBuilder.ToString();
		string text2 = text.Substring(0, match.Index);
		if (ConvertUtil._invariantCompareInfo.IsSuffix(match.Value, "/>"))
		{
			text2 += text.Substring(match.Index, text.Length - match.Index);
		}
		else
		{
			Match match2;
			if (streamReader.Peek() != -1)
			{
				long num3 = end;
				while (num3 >= 0)
				{
					num3 = Math.Max(num3 - 8192, 0L);
					int num4 = (int)(end - num3);
					stream.Seek(num3, SeekOrigin.Begin);
					char[] array = new char[num4];
					streamReader = new StreamReader(stream);
					int charCount = streamReader.ReadBlock(array, 0, num4);
					stringBuilder = new StringBuilder();
					stringBuilder.Append(array, 0, charCount);
					text = stringBuilder.ToString();
					match2 = Regex.Match(text, string.Format("(</[^>]*{0}[^>]*>)", "sheetData"));
					if (match2.Success)
					{
						break;
					}
				}
			}
			match2 = Regex.Match(text, string.Format("(</[^>]*{0}[^>]*>)", "sheetData"));
			text2 = text2 + "<sheetData/>" + text.Substring(match2.Index + match2.Length, text.Length - (match2.Index + match2.Length));
		}
		if (streamReader.Peek() > -1)
		{
			text2 += streamReader.ReadToEnd();
		}
		encoding = streamReader.CurrentEncoding;
		return text2;
	}

	private void GetBlockPos(string xml, string tag, ref int start, ref int end)
	{
		Match match = Regex.Match(xml.Substring(start), $"(<[^>]*{tag}[^>]*>)");
		if (!match.Success)
		{
			start = -1;
			end = -1;
			return;
		}
		int num = match.Index + start;
		if (match.Value.Substring(match.Value.Length - 2, 1) == "/")
		{
			end = num + match.Length;
		}
		else
		{
			Match match2 = Regex.Match(xml.Substring(start), $"(</[^>]*{tag}[^>]*>)");
			if (match2.Success)
			{
				end = match2.Index + match2.Length + start;
			}
		}
		start = num;
	}

	private bool ReadUntil(XmlReader xr, params string[] tagName)
	{
		if (xr.EOF)
		{
			return false;
		}
		while (!Array.Exists(tagName, (string tag) => ConvertUtil._invariantCompareInfo.IsSuffix(xr.LocalName, tag)))
		{
			xr.Read();
			if (xr.EOF)
			{
				return false;
			}
		}
		return ConvertUtil._invariantCompareInfo.IsSuffix(xr.LocalName, tagName[0]);
	}

	private void LoadColumns(XmlReader xr)
	{
		new List<IRangeID>();
		if (!ReadUntil(xr, "cols", "sheetData"))
		{
			return;
		}
		while (xr.Read())
		{
			if (xr.NodeType == XmlNodeType.Whitespace)
			{
				continue;
			}
			if (xr.LocalName != "col")
			{
				break;
			}
			if (xr.NodeType == XmlNodeType.Element)
			{
				int col = int.Parse(xr.GetAttribute("min"));
				ExcelColumn excelColumn = new ExcelColumn(this, col);
				excelColumn.ColumnMax = int.Parse(xr.GetAttribute("max"));
				excelColumn.Width = ((xr.GetAttribute("width") == null) ? 0.0 : double.Parse(xr.GetAttribute("width"), CultureInfo.InvariantCulture));
				excelColumn.BestFit = ((xr.GetAttribute("bestFit") != null && xr.GetAttribute("bestFit") == "1") ? true : false);
				excelColumn.Collapsed = ((xr.GetAttribute("collapsed") != null && xr.GetAttribute("collapsed") == "1") ? true : false);
				excelColumn.Phonetic = ((xr.GetAttribute("phonetic") != null && xr.GetAttribute("phonetic") == "1") ? true : false);
				excelColumn.OutlineLevel = (short)((xr.GetAttribute("outlineLevel") != null) ? int.Parse(xr.GetAttribute("outlineLevel"), CultureInfo.InvariantCulture) : 0);
				excelColumn.Hidden = ((xr.GetAttribute("hidden") != null && xr.GetAttribute("hidden") == "1") ? true : false);
				SetValueInner(0, col, excelColumn);
				if (xr.GetAttribute("style") != null && int.TryParse(xr.GetAttribute("style"), NumberStyles.Number, CultureInfo.InvariantCulture, out var result))
				{
					SetStyleInner(0, col, result);
				}
			}
		}
	}

	private static bool ReadXmlReaderUntil(XmlReader xr, string nodeText, string altNode)
	{
		do
		{
			if (xr.LocalName == nodeText || xr.LocalName == altNode)
			{
				return true;
			}
		}
		while (xr.Read());
		xr.Close();
		return false;
	}

	private void LoadHyperLinks(XmlReader xr)
	{
		if (!ReadUntil(xr, "hyperlinks", "rowBreaks", "colBreaks"))
		{
			return;
		}
		while (xr.Read() && xr.LocalName == "hyperlink")
		{
			ExcelCellBase.GetRowColFromAddress(xr.GetAttribute("ref"), out int FromRow, out int FromColumn, out int ToRow, out int ToColumn);
			ExcelHyperLink excelHyperLink = null;
			if (xr.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships") != null)
			{
				string attribute = xr.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
				Uri targetUri = Part.GetRelationship(attribute).TargetUri;
				if (targetUri.IsAbsoluteUri)
				{
					try
					{
						excelHyperLink = new ExcelHyperLink(targetUri.AbsoluteUri);
					}
					catch
					{
						excelHyperLink = new ExcelHyperLink(targetUri.OriginalString, UriKind.Absolute);
					}
				}
				else
				{
					excelHyperLink = new ExcelHyperLink(targetUri.OriginalString, UriKind.Relative);
				}
				excelHyperLink.RId = attribute;
				Part.DeleteRelationship(attribute);
			}
			else if (xr.GetAttribute("location") != null)
			{
				excelHyperLink = GetHyperlinkFromRef(xr, "location", FromRow, ToRow, FromColumn, ToColumn);
			}
			else
			{
				if (xr.GetAttribute("ref") == null)
				{
					break;
				}
				excelHyperLink = GetHyperlinkFromRef(xr, "ref", FromRow, ToRow, FromColumn, ToColumn);
			}
			string attribute2 = xr.GetAttribute("tooltip");
			if (!string.IsNullOrEmpty(attribute2))
			{
				excelHyperLink.ToolTip = attribute2;
			}
			_hyperLinks.SetValue(FromRow, FromColumn, excelHyperLink);
		}
	}

	private ExcelHyperLink GetHyperlinkFromRef(XmlReader xr, string refTag, int fromRow = 0, int toRow = 0, int fromCol = 0, int toCol = 0)
	{
		return new ExcelHyperLink(xr.GetAttribute(refTag), xr.GetAttribute("display"))
		{
			RowSpann = toRow - fromRow,
			ColSpann = toCol - fromCol
		};
	}

	private void LoadCells(XmlReader xr)
	{
		ReadUntil(xr, "sheetData", "mergeCells", "hyperlinks", "rowBreaks", "colBreaks");
		ExcelAddressBase excelAddressBase = null;
		string type = "";
		int num = 0;
		int num2 = 0;
		int num3 = 0;
		xr.Read();
		while (!xr.EOF)
		{
			while (xr.NodeType == XmlNodeType.EndElement)
			{
				xr.Read();
			}
			if (xr.LocalName == "row")
			{
				num3 = 0;
				string attribute = xr.GetAttribute("r");
				num2 = ((attribute != null) ? Convert.ToInt32(attribute) : (num2 + 1));
				if (DoAddRow(xr))
				{
					SetValueInner(num2, 0, AddRow(xr, num2));
					if (xr.GetAttribute("s") != null)
					{
						SetStyleInner(num2, 0, int.Parse(xr.GetAttribute("s"), CultureInfo.InvariantCulture));
					}
				}
				xr.Read();
				continue;
			}
			if (xr.LocalName == "c")
			{
				string attribute2 = xr.GetAttribute("r");
				if (attribute2 == null)
				{
					num3++;
					excelAddressBase = new ExcelAddressBase(num2, num3, num2, num3);
				}
				else
				{
					excelAddressBase = new ExcelAddressBase(attribute2);
					num3 = excelAddressBase._fromCol;
				}
				type = ((xr.GetAttribute("t") == null) ? "" : xr.GetAttribute("t"));
				if (xr.GetAttribute("s") != null)
				{
					num = int.Parse(xr.GetAttribute("s"));
					SetStyleInner(excelAddressBase._fromRow, excelAddressBase._fromCol, num);
				}
				else
				{
					num = 0;
				}
				xr.Read();
				continue;
			}
			if (xr.LocalName == "v")
			{
				SetValueFromXml(xr, type, num, excelAddressBase._fromRow, excelAddressBase._fromCol);
				xr.Read();
				continue;
			}
			if (xr.LocalName == "f")
			{
				string attribute3 = xr.GetAttribute("t");
				if (attribute3 == null)
				{
					_formulas.SetValue(excelAddressBase._fromRow, excelAddressBase._fromCol, xr.ReadElementContentAsString());
					SetValueInner(excelAddressBase._fromRow, excelAddressBase._fromCol, null);
				}
				else if (attribute3 == "shared")
				{
					string attribute4 = xr.GetAttribute("si");
					if (attribute4 != null)
					{
						int num4 = int.Parse(attribute4);
						_formulas.SetValue(excelAddressBase._fromRow, excelAddressBase._fromCol, num4);
						SetValueInner(excelAddressBase._fromRow, excelAddressBase._fromCol, null);
						string attribute5 = xr.GetAttribute("ref");
						string text = ConvertUtil.ExcelDecodeString(xr.ReadElementContentAsString());
						if (text != "")
						{
							_sharedFormulas.Add(num4, new Formulas(SourceCodeTokenizer.Default)
							{
								Index = num4,
								Formula = text,
								Address = attribute5,
								StartRow = excelAddressBase._fromRow,
								StartCol = excelAddressBase._fromCol
							});
						}
					}
					else
					{
						xr.Read();
					}
				}
				else if (attribute3 == "array")
				{
					string attribute6 = xr.GetAttribute("ref");
					string formula = xr.ReadElementContentAsString();
					int maxShareFunctionIndex = GetMaxShareFunctionIndex(isArray: true);
					_formulas.SetValue(excelAddressBase._fromRow, excelAddressBase._fromCol, maxShareFunctionIndex);
					SetValueInner(excelAddressBase._fromRow, excelAddressBase._fromCol, null);
					_sharedFormulas.Add(maxShareFunctionIndex, new Formulas(SourceCodeTokenizer.Default)
					{
						Index = maxShareFunctionIndex,
						Formula = formula,
						Address = attribute6,
						StartRow = excelAddressBase._fromRow,
						StartCol = excelAddressBase._fromCol,
						IsArray = true
					});
					_flags.SetFlagValue(excelAddressBase._fromRow, excelAddressBase._fromCol, value: true, CellFlags.ArrayFormula);
				}
				else
				{
					xr.Read();
				}
				continue;
			}
			if (!(xr.LocalName == "is"))
			{
				break;
			}
			xr.Read();
			if (xr.LocalName == "t")
			{
				SetValueInner(excelAddressBase._fromRow, excelAddressBase._fromCol, ConvertUtil.ExcelDecodeString(xr.ReadElementContentAsString()));
				continue;
			}
			if (xr.LocalName == "r")
			{
				string text2 = xr.ReadOuterXml();
				while (xr.LocalName == "r")
				{
					text2 += xr.ReadOuterXml();
				}
				SetValueInner(excelAddressBase._fromRow, excelAddressBase._fromCol, text2);
			}
			else
			{
				SetValueInner(excelAddressBase._fromRow, excelAddressBase._fromCol, xr.ReadOuterXml());
			}
			_flags.SetFlagValue(excelAddressBase._fromRow, excelAddressBase._fromCol, value: true, CellFlags.RichText);
		}
	}

	private bool DoAddRow(XmlReader xr)
	{
		int num = ((xr.GetAttribute("r") != null) ? 1 : 0);
		if (xr.GetAttribute("spans") != null)
		{
			num++;
		}
		return xr.AttributeCount > num;
	}

	private void LoadMergeCells(XmlReader xr)
	{
		if (!ReadUntil(xr, "mergeCells", "hyperlinks", "rowBreaks", "colBreaks") || xr.EOF)
		{
			return;
		}
		while (xr.Read() && !(xr.LocalName != "mergeCell"))
		{
			if (xr.NodeType == XmlNodeType.Element)
			{
				string attribute = xr.GetAttribute("ref");
				_mergedCells.Add(new ExcelAddress(attribute), doValidate: false);
			}
		}
	}

	private void UpdateMergedCells(StreamWriter sw)
	{
		sw.Write("<mergeCells>");
		foreach (string mergedCell in _mergedCells)
		{
			sw.Write("<mergeCell ref=\"{0}\" />", mergedCell);
		}
		sw.Write("</mergeCells>");
	}

	private RowInternal AddRow(XmlReader xr, int row)
	{
		return new RowInternal
		{
			Collapsed = ((xr.GetAttribute("collapsed") != null && xr.GetAttribute("collapsed") == "1") ? true : false),
			OutlineLevel = (short)((xr.GetAttribute("outlineLevel") != null) ? short.Parse(xr.GetAttribute("outlineLevel"), CultureInfo.InvariantCulture) : 0),
			Height = ((xr.GetAttribute("ht") == null) ? (-1.0) : double.Parse(xr.GetAttribute("ht"), CultureInfo.InvariantCulture)),
			Hidden = ((xr.GetAttribute("hidden") != null && xr.GetAttribute("hidden") == "1") ? true : false),
			Phonetic = ((xr.GetAttribute("ph") != null && xr.GetAttribute("ph") == "1") ? true : false),
			CustomHeight = (xr.GetAttribute("customHeight") != null && xr.GetAttribute("customHeight") == "1")
		};
	}

	private void SetValueFromXml(XmlReader xr, string type, int styleID, int row, int col)
	{
		switch (type)
		{
		case "s":
		{
			int index = xr.ReadElementContentAsInt();
			SetValueInner(row, col, _package.Workbook._sharedStringsList[index].Text);
			if (_package.Workbook._sharedStringsList[index].isRichText)
			{
				_flags.SetFlagValue(row, col, value: true, CellFlags.RichText);
			}
			return;
		}
		case "str":
			SetValueInner(row, col, ConvertUtil.ExcelDecodeString(xr.ReadElementContentAsString()));
			return;
		case "b":
			SetValueInner(row, col, xr.ReadElementContentAsString() != "0");
			return;
		case "e":
			SetValueInner(row, col, GetErrorType(xr.ReadElementContentAsString()));
			return;
		}
		string text = xr.ReadElementContentAsString();
		int numberFormatId = Workbook.Styles.CellXfs[styleID].NumberFormatId;
		double result2;
		if ((numberFormatId >= 14 && numberFormatId <= 22) || (numberFormatId >= 45 && numberFormatId <= 47))
		{
			if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out var result))
			{
				if (Workbook.Date1904)
				{
					result += 1462.0;
				}
				if (result >= -657435.0 && result < 2958465.9999999)
				{
					SetValueInner(row, col, DateTime.FromOADate(result));
				}
				else
				{
					SetValueInner(row, col, result);
				}
			}
			else
			{
				SetValueInner(row, col, text);
			}
		}
		else if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out result2))
		{
			SetValueInner(row, col, result2);
		}
		else
		{
			SetValueInner(row, col, double.NaN);
		}
	}

	private object GetErrorType(string v)
	{
		return ExcelErrorValue.Parse(ConvertUtil._invariantTextInfo.ToUpper(v));
	}

	public ExcelRow Row(int row)
	{
		CheckSheetType();
		if (row < 1 || row > 1048576)
		{
			throw new ArgumentException("Row number out of bounds");
		}
		return new ExcelRow(this, row);
	}

	public ExcelColumn Column(int col)
	{
		CheckSheetType();
		if (col < 1 || col > 16384)
		{
			throw new ArgumentException("Column number out of bounds");
		}
		ExcelColumn excelColumn = GetValueInner(0, col) as ExcelColumn;
		if (excelColumn != null)
		{
			if (excelColumn.ColumnMin != excelColumn.ColumnMax)
			{
				int columnMax = excelColumn.ColumnMax;
				excelColumn.ColumnMax = col;
				CopyColumn(excelColumn, col + 1, columnMax);
			}
		}
		else
		{
			int row = 0;
			int col2 = col;
			if (_values.PrevCell(ref row, ref col2))
			{
				excelColumn = GetValueInner(0, col2) as ExcelColumn;
				int columnMax2 = excelColumn.ColumnMax;
				if (columnMax2 >= col)
				{
					excelColumn.ColumnMax = col - 1;
					if (columnMax2 > col)
					{
						CopyColumn(excelColumn, col + 1, columnMax2);
					}
					return CopyColumn(excelColumn, col, col);
				}
			}
			excelColumn = new ExcelColumn(this, col);
			SetValueInner(0, col, excelColumn);
		}
		return excelColumn;
	}

	public override string ToString()
	{
		return Name;
	}

	internal ExcelColumn CopyColumn(ExcelColumn c, int col, int maxCol)
	{
		ExcelColumn excelColumn = new ExcelColumn(this, col);
		excelColumn.ColumnMax = ((maxCol < 16384) ? maxCol : 16384);
		if (c.StyleName != "")
		{
			excelColumn.StyleName = c.StyleName;
		}
		else
		{
			excelColumn.StyleID = c.StyleID;
		}
		excelColumn.OutlineLevel = c.OutlineLevel;
		excelColumn.Phonetic = c.Phonetic;
		excelColumn.BestFit = c.BestFit;
		SetValueInner(0, col, excelColumn);
		excelColumn._width = c._width;
		excelColumn._hidden = c._hidden;
		return excelColumn;
	}

	public void Select()
	{
		View.TabSelected = true;
	}

	public void Select(string Address)
	{
		Select(Address, SelectSheet: true);
	}

	public void Select(string Address, bool SelectSheet)
	{
		CheckSheetType();
		ExcelCellBase.GetRowColFromAddress(Address, out int FromRow, out int FromColumn, out int _, out int _);
		if (SelectSheet)
		{
			View.TabSelected = true;
		}
		View.SelectedRange = Address;
		View.ActiveCell = ExcelCellBase.GetAddress(FromRow, FromColumn);
	}

	public void Select(ExcelAddress Address)
	{
		CheckSheetType();
		Select(Address, SelectSheet: true);
	}

	public void Select(ExcelAddress Address, bool SelectSheet)
	{
		CheckSheetType();
		if (SelectSheet)
		{
			View.TabSelected = true;
		}
		string text = ExcelCellBase.GetAddress(Address.Start.Row, Address.Start.Column) + ":" + ExcelCellBase.GetAddress(Address.End.Row, Address.End.Column);
		if (Address.Addresses != null)
		{
			foreach (ExcelAddress address in Address.Addresses)
			{
				text = text + " " + ExcelCellBase.GetAddress(address.Start.Row, address.Start.Column) + ":" + ExcelCellBase.GetAddress(address.End.Row, address.End.Column);
			}
		}
		View.SelectedRange = text;
		View.ActiveCell = ExcelCellBase.GetAddress(Address.Start.Row, Address.Start.Column);
	}

	public void InsertRow(int rowFrom, int rows)
	{
		InsertRow(rowFrom, rows, 0);
	}

	public void InsertRow(int rowFrom, int rows, int copyStylesFromRow)
	{
		CheckSheetType();
		ExcelAddressBase dimension = Dimension;
		if (rowFrom < 1)
		{
			throw new ArgumentOutOfRangeException("rowFrom can't be lesser that 1");
		}
		if (dimension != null && dimension.End.Row > rowFrom && dimension.End.Row + rows > 1048576)
		{
			throw new ArgumentOutOfRangeException("Can't insert. Rows will be shifted outside the boundries of the worksheet.");
		}
		lock (this)
		{
			_values.Insert(rowFrom, 0, rows, 0);
			_formulas.Insert(rowFrom, 0, rows, 0);
			_commentsStore.Insert(rowFrom, 0, rows, 0);
			_hyperLinks.Insert(rowFrom, 0, rows, 0);
			_flags.Insert(rowFrom, 0, rows, 0);
			Comments.Insert(rowFrom, 0, rows, 0);
			_names.Insert(rowFrom, 0, rows, 0);
			Workbook.Names.Insert(rowFrom, 0, rows, 0, (ExcelNamedRange n) => n.Worksheet == this);
			foreach (Formulas value in _sharedFormulas.Values)
			{
				if (value.StartRow >= rowFrom)
				{
					value.StartRow += rows;
				}
				ExcelAddressBase excelAddressBase = new ExcelAddressBase(value.Address);
				if (excelAddressBase._fromRow >= rowFrom)
				{
					excelAddressBase._fromRow += rows;
					excelAddressBase._toRow += rows;
				}
				else if (excelAddressBase._toRow >= rowFrom)
				{
					excelAddressBase._toRow += rows;
				}
				value.Address = ExcelCellBase.GetAddress(excelAddressBase._fromRow, excelAddressBase._fromCol, excelAddressBase._toRow, excelAddressBase._toCol);
				value.Formula = ExcelCellBase.UpdateFormulaReferences(value.Formula, rows, 0, rowFrom, 0, Name, Name);
			}
			CellsStoreEnumerator<object> cellsStoreEnumerator = new CellsStoreEnumerator<object>(_formulas);
			while (cellsStoreEnumerator.Next())
			{
				if (cellsStoreEnumerator.Value is string)
				{
					cellsStoreEnumerator.Value = ExcelCellBase.UpdateFormulaReferences(cellsStoreEnumerator.Value.ToString(), rows, 0, rowFrom, 0, Name, Name);
				}
			}
			FixMergedCellsRow(rowFrom, rows, delete: false);
			if (copyStylesFromRow > 0)
			{
				CellsStoreEnumerator<ExcelCoreValue> cellsStoreEnumerator2 = new CellsStoreEnumerator<ExcelCoreValue>(_values, copyStylesFromRow, 0, copyStylesFromRow, 16384);
				while (cellsStoreEnumerator2.Next())
				{
					if (cellsStoreEnumerator2.Value._styleId != 0)
					{
						for (int i = 0; i < rows; i++)
						{
							SetStyleInner(rowFrom + i, cellsStoreEnumerator2.Column, cellsStoreEnumerator2.Value._styleId);
						}
					}
				}
				int outlineLevel = Row(copyStylesFromRow + rows).OutlineLevel;
				for (int j = 0; j < rows; j++)
				{
					Row(rowFrom + j).OutlineLevel = outlineLevel;
				}
			}
			foreach (ExcelTable table in Tables)
			{
				table.Address = table.Address.AddRow(rowFrom, rows);
			}
			foreach (ExcelPivotTable pivotTable in PivotTables)
			{
				pivotTable.Address = pivotTable.Address.AddRow(rowFrom, rows);
				pivotTable.CacheDefinition.SourceRange.Address = pivotTable.CacheDefinition.SourceRange.AddRow(rowFrom, rows).Address;
			}
			foreach (ExcelDataValidation item in (IEnumerable<IExcelDataValidation>)DataValidations)
			{
				ExcelAddress address = item.Address;
				string address2 = address.AddRow(rowFrom, rows).Address;
				if (address.Address != address2)
				{
					item.SetAddress(address2);
				}
			}
		}
		foreach (ExcelWorksheet item2 in Workbook.Worksheets.Where((ExcelWorksheet sheet) => sheet != this))
		{
			item2.UpdateCrossSheetReferences(Name, rowFrom, rows, 0, 0);
		}
	}

	public void InsertColumn(int columnFrom, int columns)
	{
		InsertColumn(columnFrom, columns, 0);
	}

	public void InsertColumn(int columnFrom, int columns, int copyStylesFromColumn)
	{
		CheckSheetType();
		ExcelAddressBase dimension = Dimension;
		if (columnFrom < 1)
		{
			throw new ArgumentOutOfRangeException("columnFrom can't be lesser that 1");
		}
		if (dimension != null && dimension.End.Column > columnFrom && dimension.End.Column + columns > 16384)
		{
			throw new ArgumentOutOfRangeException("Can't insert. Columns will be shifted outside the boundries of the worksheet.");
		}
		lock (this)
		{
			_values.Insert(0, columnFrom, 0, columns);
			_formulas.Insert(0, columnFrom, 0, columns);
			_commentsStore.Insert(0, columnFrom, 0, columns);
			_hyperLinks.Insert(0, columnFrom, 0, columns);
			_flags.Insert(0, columnFrom, 0, columns);
			_names.Insert(0, columnFrom, 0, columns);
			Comments.Insert(0, columnFrom, 0, columns);
			Workbook.Names.Insert(0, columnFrom, 0, columns, (ExcelNamedRange n) => n.Worksheet == this);
			foreach (Formulas value2 in _sharedFormulas.Values)
			{
				if (value2.StartCol >= columnFrom)
				{
					value2.StartCol += columns;
				}
				ExcelAddressBase excelAddressBase = new ExcelAddressBase(value2.Address);
				if (excelAddressBase._fromCol >= columnFrom)
				{
					excelAddressBase._fromCol += columns;
					excelAddressBase._toCol += columns;
				}
				else if (excelAddressBase._toCol >= columnFrom)
				{
					excelAddressBase._toCol += columns;
				}
				value2.Address = ExcelCellBase.GetAddress(excelAddressBase._fromRow, excelAddressBase._fromCol, excelAddressBase._toRow, excelAddressBase._toCol);
				value2.Formula = ExcelCellBase.UpdateFormulaReferences(value2.Formula, 0, columns, 0, columnFrom, Name, Name);
			}
			CellsStoreEnumerator<object> cellsStoreEnumerator = new CellsStoreEnumerator<object>(_formulas);
			while (cellsStoreEnumerator.Next())
			{
				if (cellsStoreEnumerator.Value is string)
				{
					cellsStoreEnumerator.Value = ExcelCellBase.UpdateFormulaReferences(cellsStoreEnumerator.Value.ToString(), 0, columns, 0, columnFrom, Name, Name);
				}
			}
			FixMergedCellsColumn(columnFrom, columns, delete: false);
			CellsStoreEnumerator<ExcelCoreValue> cellsStoreEnumerator2 = new CellsStoreEnumerator<ExcelCoreValue>(_values, 0, 1, 0, 16384);
			List<ExcelColumn> list = new List<ExcelColumn>();
			foreach (ExcelCoreValue item in cellsStoreEnumerator2)
			{
				object value = item._value;
				if (value is ExcelColumn)
				{
					list.Add((ExcelColumn)value);
				}
			}
			for (int num = list.Count - 1; num >= 0; num--)
			{
				ExcelColumn excelColumn = list[num];
				if (excelColumn._columnMin >= columnFrom)
				{
					if (excelColumn._columnMin + columns <= 16384)
					{
						excelColumn._columnMin += columns;
					}
					else
					{
						excelColumn._columnMin = 16384;
					}
					if (excelColumn._columnMax + columns <= 16384)
					{
						excelColumn._columnMax += columns;
					}
					else
					{
						excelColumn._columnMax = 16384;
					}
				}
				else if (excelColumn._columnMax >= columnFrom)
				{
					int num2 = excelColumn._columnMax - columnFrom;
					excelColumn._columnMax = columnFrom - 1;
					CopyColumn(excelColumn, columnFrom + columns, columnFrom + columns + num2);
				}
			}
			if (copyStylesFromColumn > 0)
			{
				if (copyStylesFromColumn >= columnFrom)
				{
					copyStylesFromColumn += columns;
				}
				List<int[]> list2 = new List<int[]>();
				CellsStoreEnumerator<ExcelCoreValue> cellsStoreEnumerator3 = new CellsStoreEnumerator<ExcelCoreValue>(_values, 0, copyStylesFromColumn, 1048576, copyStylesFromColumn);
				lock (cellsStoreEnumerator3)
				{
					while (cellsStoreEnumerator3.Next())
					{
						if (cellsStoreEnumerator3.Value._styleId != 0)
						{
							list2.Add(new int[2]
							{
								cellsStoreEnumerator3.Row,
								cellsStoreEnumerator3.Value._styleId
							});
						}
					}
				}
				foreach (int[] item2 in list2)
				{
					for (int i = 0; i < columns; i++)
					{
						if (item2[0] == 0)
						{
							Column(columnFrom + i).StyleID = item2[1];
						}
						else
						{
							SetStyleInner(item2[0], columnFrom + i, item2[1]);
						}
					}
				}
				int outlineLevel = Column(copyStylesFromColumn).OutlineLevel;
				for (int j = 0; j < columns; j++)
				{
					Column(columnFrom + j).OutlineLevel = outlineLevel;
				}
			}
			foreach (ExcelTable table in Tables)
			{
				if (columnFrom > table.Address.Start.Column && columnFrom <= table.Address.End.Column)
				{
					InsertTableColumns(columnFrom, columns, table);
				}
				table.Address = table.Address.AddColumn(columnFrom, columns);
			}
			foreach (ExcelPivotTable pivotTable in PivotTables)
			{
				if (columnFrom <= pivotTable.Address.End.Column)
				{
					pivotTable.Address = pivotTable.Address.AddColumn(columnFrom, columns);
				}
				if (columnFrom <= pivotTable.CacheDefinition.SourceRange.End.Column && pivotTable.CacheDefinition.CacheSource == eSourceType.Worksheet)
				{
					pivotTable.CacheDefinition.SourceRange.Address = pivotTable.CacheDefinition.SourceRange.AddColumn(columnFrom, columns).Address;
				}
			}
			foreach (ExcelDataValidation item3 in (IEnumerable<IExcelDataValidation>)DataValidations)
			{
				ExcelAddress address = item3.Address;
				string address2 = address.AddColumn(columnFrom, columns).Address;
				if (address.Address != address2)
				{
					item3.SetAddress(address2);
				}
			}
			foreach (ExcelWorksheet item4 in Workbook.Worksheets.Where((ExcelWorksheet sheet) => sheet != this))
			{
				item4.UpdateCrossSheetReferences(Name, 0, 0, columnFrom, columns);
			}
		}
	}

	private static void InsertTableColumns(int columnFrom, int columns, ExcelTable tbl)
	{
		XmlNode parentNode = tbl.Columns[0].TopNode.ParentNode;
		int num = columnFrom - tbl.Address.Start.Column - 1;
		XmlNode refChild = parentNode.ChildNodes[num];
		num += 2;
		for (int i = 0; i < columns; i++)
		{
			string uniqueName = tbl.Columns.GetUniqueName($"Column{num++.ToString(CultureInfo.InvariantCulture)}");
			XmlElement xmlElement = (XmlElement)tbl.TableXml.CreateNode(XmlNodeType.Element, "tableColumn", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
			xmlElement.SetAttribute("id", (tbl.Columns.Count + i + 1).ToString(CultureInfo.InvariantCulture));
			xmlElement.SetAttribute("name", uniqueName);
			refChild = parentNode.InsertAfter(xmlElement, refChild);
		}
		tbl._cols = new ExcelTableColumnCollection(tbl);
	}

	private void FixMergedCellsRow(int row, int rows, bool delete)
	{
		if (delete)
		{
			_mergedCells._cells.Delete(row, 0, rows, 0);
		}
		else
		{
			_mergedCells._cells.Insert(row, 0, rows, 0);
		}
		List<int> list = new List<int>();
		for (int i = 0; i < _mergedCells.Count; i++)
		{
			if (string.IsNullOrEmpty(_mergedCells[i]))
			{
				continue;
			}
			ExcelAddressBase excelAddressBase = new ExcelAddressBase(_mergedCells[i]);
			ExcelAddressBase excelAddressBase2;
			if (delete)
			{
				excelAddressBase2 = excelAddressBase.DeleteRow(row, rows);
				if (excelAddressBase2 == null)
				{
					list.Add(i);
					continue;
				}
			}
			else
			{
				excelAddressBase2 = excelAddressBase.AddRow(row, rows);
				if (excelAddressBase2.Address != excelAddressBase.Address)
				{
					_mergedCells.SetIndex(excelAddressBase2, i);
				}
			}
			if (excelAddressBase2.Address != excelAddressBase.Address)
			{
				_mergedCells.List[i] = excelAddressBase2._address;
			}
		}
		for (int num = list.Count - 1; num >= 0; num--)
		{
			_mergedCells.List.RemoveAt(list[num]);
		}
	}

	private void FixMergedCellsColumn(int column, int columns, bool delete)
	{
		if (delete)
		{
			_mergedCells._cells.Delete(0, column, 0, columns);
		}
		else
		{
			_mergedCells._cells.Insert(0, column, 0, columns);
		}
		List<int> list = new List<int>();
		for (int i = 0; i < _mergedCells.Count; i++)
		{
			if (string.IsNullOrEmpty(_mergedCells[i]))
			{
				continue;
			}
			ExcelAddressBase excelAddressBase = new ExcelAddressBase(_mergedCells[i]);
			ExcelAddressBase excelAddressBase2;
			if (delete)
			{
				excelAddressBase2 = excelAddressBase.DeleteColumn(column, columns);
				if (excelAddressBase2 == null)
				{
					list.Add(i);
					continue;
				}
			}
			else
			{
				excelAddressBase2 = excelAddressBase.AddColumn(column, columns);
				if (excelAddressBase2.Address != excelAddressBase.Address)
				{
					_mergedCells.SetIndex(excelAddressBase2, i);
				}
			}
			if (excelAddressBase2.Address != excelAddressBase.Address)
			{
				_mergedCells.List[i] = excelAddressBase2._address;
			}
		}
		for (int num = list.Count - 1; num >= 0; num--)
		{
			_mergedCells.List.RemoveAt(list[num]);
		}
	}

	private void FixSharedFormulasRows(int position, int rows)
	{
		List<Formulas> list = new List<Formulas>();
		List<Formulas> list2 = new List<Formulas>();
		foreach (int key in _sharedFormulas.Keys)
		{
			Formulas formulas = _sharedFormulas[key];
			ExcelCellBase.GetRowColFromAddress(formulas.Address, out int FromRow, out int FromColumn, out int ToRow, out int ToColumn);
			if (position >= FromRow && position + Math.Abs(rows) <= ToRow)
			{
				if (rows > 0)
				{
					formulas.Address = ExcelCellBase.GetAddress(FromRow, FromColumn) + ":" + ExcelCellBase.GetAddress(position - 1, ToColumn);
					if (ToRow != FromRow)
					{
						Formulas formulas2 = new Formulas(SourceCodeTokenizer.Default);
						formulas2.StartCol = formulas.StartCol;
						formulas2.StartRow = position + rows;
						formulas2.Address = ExcelCellBase.GetAddress(position + rows, FromColumn) + ":" + ExcelCellBase.GetAddress(ToRow + rows, ToColumn);
						formulas2.Formula = ExcelCellBase.TranslateFromR1C1(ExcelCellBase.TranslateToR1C1(formulas.Formula, formulas.StartRow, formulas.StartCol), position, formulas.StartCol);
						list.Add(formulas2);
					}
				}
				else if (FromRow - rows < ToRow)
				{
					formulas.Address = ExcelCellBase.GetAddress(FromRow, FromColumn, ToRow + rows, ToColumn);
				}
				else
				{
					formulas.Address = ExcelCellBase.GetAddress(FromRow, FromColumn) + ":" + ExcelCellBase.GetAddress(ToRow + rows, ToColumn);
				}
			}
			else
			{
				if (position > ToRow)
				{
					continue;
				}
				if (rows > 0)
				{
					formulas.StartRow += rows;
					formulas.Address = ExcelCellBase.GetAddress(FromRow + rows, FromColumn) + ":" + ExcelCellBase.GetAddress(ToRow + rows, ToColumn);
					continue;
				}
				if (position <= FromRow && position + Math.Abs(rows) > ToRow)
				{
					list2.Add(formulas);
					continue;
				}
				ToRow = ((ToRow + rows < position - 1) ? (position - 1) : (ToRow + rows));
				if (position <= FromRow)
				{
					FromRow = ((FromRow + rows < position) ? position : (FromRow + rows));
				}
				formulas.Address = ExcelCellBase.GetAddress(FromRow, FromColumn, ToRow, ToColumn);
				Cells[formulas.Address].SetSharedFormulaID(formulas.Index);
			}
		}
		AddFormulas(list, position, rows);
		foreach (Formulas item in list2)
		{
			_sharedFormulas.Remove(item.Index);
		}
		list = new List<Formulas>();
		foreach (int key2 in _sharedFormulas.Keys)
		{
			Formulas formula = _sharedFormulas[key2];
			UpdateSharedFormulaRow(ref formula, position, rows, ref list);
		}
		AddFormulas(list, position, rows);
	}

	private void AddFormulas(List<Formulas> added, int position, int rows)
	{
		foreach (Formulas item in added)
		{
			item.Index = GetMaxShareFunctionIndex(isArray: false);
			_sharedFormulas.Add(item.Index, item);
			Cells[item.Address].SetSharedFormulaID(item.Index);
		}
	}

	private void UpdateSharedFormulaRow(ref Formulas formula, int startRow, int rows, ref List<Formulas> newFormulas)
	{
		int count = newFormulas.Count;
		ExcelCellBase.GetRowColFromAddress(formula.Address, out int FromRow, out int FromColumn, out int ToRow, out int ToColumn);
		string text;
		if (rows > 0 || FromRow <= startRow)
		{
			text = ExcelCellBase.TranslateToR1C1(formula.Formula, formula.StartRow, formula.StartCol);
			formula.Formula = ExcelCellBase.TranslateFromR1C1(text, FromRow, formula.StartCol);
		}
		else
		{
			text = ExcelCellBase.TranslateToR1C1(formula.Formula, formula.StartRow - rows, formula.StartCol);
			formula.Formula = ExcelCellBase.TranslateFromR1C1(text, formula.StartRow, formula.StartCol);
		}
		string text2 = text;
		for (int i = FromRow; i <= ToRow; i++)
		{
			for (int j = FromColumn; j <= ToColumn; j++)
			{
				string text3;
				string text4;
				if (rows > 0 || i < startRow)
				{
					text3 = ExcelCellBase.UpdateFormulaReferences(ExcelCellBase.TranslateFromR1C1(text, i, j), rows, 0, startRow, 0, Name, Name);
					text4 = ExcelCellBase.TranslateToR1C1(text3, i, j);
				}
				else
				{
					text3 = ExcelCellBase.UpdateFormulaReferences(ExcelCellBase.TranslateFromR1C1(text, i - rows, j), rows, 0, startRow, 0, Name, Name);
					text4 = ExcelCellBase.TranslateToR1C1(text3, i, j);
				}
				if (!(text4 != text2))
				{
					continue;
				}
				if (i == FromRow && j == FromColumn)
				{
					formula.Formula = text3;
					continue;
				}
				if (newFormulas.Count == count)
				{
					formula.Address = ExcelCellBase.GetAddress(formula.StartRow, formula.StartCol, i - 1, j);
				}
				else
				{
					newFormulas[newFormulas.Count - 1].Address = ExcelCellBase.GetAddress(newFormulas[newFormulas.Count - 1].StartRow, newFormulas[newFormulas.Count - 1].StartCol, i - 1, j);
				}
				Formulas formulas = new Formulas(SourceCodeTokenizer.Default);
				formulas.Formula = text3;
				formulas.StartRow = i;
				formulas.StartCol = j;
				newFormulas.Add(formulas);
				text2 = text4;
			}
		}
		if (rows < 0 && formula.StartRow > startRow)
		{
			if (formula.StartRow + rows < startRow)
			{
				formula.StartRow = startRow;
			}
			else
			{
				formula.StartRow += rows;
			}
		}
		if (newFormulas.Count > count)
		{
			newFormulas[newFormulas.Count - 1].Address = ExcelCellBase.GetAddress(newFormulas[newFormulas.Count - 1].StartRow, newFormulas[newFormulas.Count - 1].StartCol, ToRow, ToColumn);
		}
	}

	public void DeleteRow(int row)
	{
		DeleteRow(row, 1);
	}

	public void DeleteRow(int rowFrom, int rows)
	{
		CheckSheetType();
		if (rowFrom < 1 || rowFrom + rows > 1048576)
		{
			throw new ArgumentException("Row out of range. Spans from 1 to " + 1048576.ToString(CultureInfo.InvariantCulture));
		}
		lock (this)
		{
			_values.Delete(rowFrom, 0, rows, 16384);
			_formulas.Delete(rowFrom, 0, rows, 16384);
			_flags.Delete(rowFrom, 0, rows, 16384);
			_commentsStore.Delete(rowFrom, 0, rows, 16384);
			_hyperLinks.Delete(rowFrom, 0, rows, 16384);
			_names.Delete(rowFrom, 0, rows, 16384);
			Comments.Delete(rowFrom, 0, rows, 16384);
			Workbook.Names.Delete(rowFrom, 0, rows, 16384, (ExcelNamedRange n) => n.Worksheet == this);
			AdjustFormulasRow(rowFrom, rows);
			FixMergedCellsRow(rowFrom, rows, delete: true);
			foreach (ExcelTable table in Tables)
			{
				table.Address = table.Address.DeleteRow(rowFrom, rows);
			}
			foreach (ExcelPivotTable pivotTable in PivotTables)
			{
				if (pivotTable.Address.Start.Row > rowFrom + rows)
				{
					pivotTable.Address = pivotTable.Address.DeleteRow(rowFrom, rows);
				}
			}
			foreach (ExcelDataValidation item in (IEnumerable<IExcelDataValidation>)DataValidations)
			{
				ExcelAddress address = item.Address;
				if (address.Start.Row > rowFrom + rows)
				{
					string address2 = address.DeleteRow(rowFrom, rows).Address;
					if (address.Address != address2)
					{
						item.SetAddress(address2);
					}
				}
			}
		}
	}

	public void DeleteColumn(int column)
	{
		DeleteColumn(column, 1);
	}

	public void DeleteColumn(int columnFrom, int columns)
	{
		if (columnFrom < 1 || columnFrom + columns > 16384)
		{
			throw new ArgumentException("Column out of range. Spans from 1 to " + 16384.ToString(CultureInfo.InvariantCulture));
		}
		lock (this)
		{
			ExcelColumn excelColumn = GetValueInner(0, columnFrom) as ExcelColumn;
			if (excelColumn == null)
			{
				int row = 0;
				int col = columnFrom;
				if (_values.PrevCell(ref row, ref col))
				{
					excelColumn = GetValueInner(0, col) as ExcelColumn;
					if (excelColumn._columnMax >= columnFrom)
					{
						excelColumn.ColumnMax = columnFrom - 1;
					}
				}
			}
			_values.Delete(0, columnFrom, 1048576, columns);
			_formulas.Delete(0, columnFrom, 1048576, columns);
			_flags.Delete(0, columnFrom, 1048576, columns);
			_commentsStore.Delete(0, columnFrom, 1048576, columns);
			_hyperLinks.Delete(0, columnFrom, 1048576, columns);
			_names.Delete(0, columnFrom, 1048576, columns);
			Comments.Delete(0, columnFrom, 0, columns);
			Workbook.Names.Delete(0, columnFrom, 1048576, columns, (ExcelNamedRange n) => n.Worksheet == this);
			AdjustFormulasColumn(columnFrom, columns);
			FixMergedCellsColumn(columnFrom, columns, delete: true);
			foreach (ExcelCoreValue item in new CellsStoreEnumerator<ExcelCoreValue>(_values, 0, columnFrom, 0, 16384))
			{
				object value = item._value;
				if (value is ExcelColumn)
				{
					ExcelColumn excelColumn2 = (ExcelColumn)value;
					if (excelColumn2._columnMin >= columnFrom)
					{
						excelColumn2._columnMin -= columns;
						excelColumn2._columnMax -= columns;
					}
				}
			}
			foreach (ExcelTable table in Tables)
			{
				if (columnFrom >= table.Address.Start.Column && columnFrom <= table.Address.End.Column)
				{
					XmlNode parentNode = table.Columns[0].TopNode.ParentNode;
					int num = columnFrom - table.Address.Start.Column;
					for (int i = 0; i < columns; i++)
					{
						if (parentNode.ChildNodes.Count > num)
						{
							parentNode.RemoveChild(parentNode.ChildNodes[num]);
						}
					}
					table._cols = new ExcelTableColumnCollection(table);
				}
				table.Address = table.Address.DeleteColumn(columnFrom, columns);
				foreach (ExcelPivotTable pivotTable in PivotTables)
				{
					if (pivotTable.Address.Start.Column > columnFrom + columns)
					{
						pivotTable.Address = pivotTable.Address.DeleteColumn(columnFrom, columns);
					}
					if (pivotTable.CacheDefinition.SourceRange.Start.Column > columnFrom + columns)
					{
						pivotTable.CacheDefinition.SourceRange.Address = pivotTable.CacheDefinition.SourceRange.DeleteColumn(columnFrom, columns).Address;
					}
				}
			}
			foreach (ExcelDataValidation item2 in (IEnumerable<IExcelDataValidation>)DataValidations)
			{
				ExcelAddress address = item2.Address;
				if (address.Start.Column > columnFrom + columns)
				{
					string address2 = address.DeleteColumn(columnFrom, columns).Address;
					if (address.Address != address2)
					{
						item2.SetAddress(address2);
					}
				}
			}
		}
	}

	internal void AdjustFormulasRow(int rowFrom, int rows)
	{
		List<int> list = new List<int>();
		foreach (Formulas value in _sharedFormulas.Values)
		{
			ExcelAddressBase excelAddressBase = new ExcelAddress(value.Address).DeleteRow(rowFrom, rows);
			if (excelAddressBase == null)
			{
				list.Add(value.Index);
				continue;
			}
			value.Address = excelAddressBase.Address;
			if (value.StartRow > rowFrom)
			{
				int num = Math.Min(value.StartRow - rowFrom, rows);
				value.Formula = ExcelCellBase.UpdateFormulaReferences(value.Formula, -num, 0, rowFrom, 0, Name, Name);
				value.StartRow -= num;
			}
		}
		foreach (int item in list)
		{
			_sharedFormulas.Remove(item);
		}
		list = null;
		CellsStoreEnumerator<object> cellsStoreEnumerator = new CellsStoreEnumerator<object>(_formulas, 1, 1, 1048576, 16384);
		while (cellsStoreEnumerator.Next())
		{
			if (cellsStoreEnumerator.Value is string)
			{
				cellsStoreEnumerator.Value = ExcelCellBase.UpdateFormulaReferences(cellsStoreEnumerator.Value.ToString(), -rows, 0, rowFrom, 0, Name, Name);
			}
		}
	}

	internal void AdjustFormulasColumn(int columnFrom, int columns)
	{
		List<int> list = new List<int>();
		foreach (Formulas value in _sharedFormulas.Values)
		{
			ExcelAddressBase excelAddressBase = new ExcelAddress(value.Address).DeleteColumn(columnFrom, columns);
			if (excelAddressBase == null)
			{
				list.Add(value.Index);
				continue;
			}
			value.Address = excelAddressBase.Address;
			if (value.StartCol > columnFrom)
			{
				int num = Math.Min(value.StartCol - columnFrom, columns);
				value.Formula = ExcelCellBase.UpdateFormulaReferences(value.Formula, 0, -num, 0, 1, Name, Name);
				value.StartCol -= num;
			}
		}
		foreach (int item in list)
		{
			_sharedFormulas.Remove(item);
		}
		list = null;
		CellsStoreEnumerator<object> cellsStoreEnumerator = new CellsStoreEnumerator<object>(_formulas, 1, 1, 1048576, 16384);
		while (cellsStoreEnumerator.Next())
		{
			if (cellsStoreEnumerator.Value is string)
			{
				cellsStoreEnumerator.Value = ExcelCellBase.UpdateFormulaReferences(cellsStoreEnumerator.Value.ToString(), 0, -columns, 0, columnFrom, Name, Name);
			}
		}
	}

	public void DeleteRow(int rowFrom, int rows, bool shiftOtherRowsUp)
	{
		DeleteRow(rowFrom, rows);
	}

	public object GetValue(int Row, int Column)
	{
		CheckSheetType();
		object valueInner = GetValueInner(Row, Column);
		if (valueInner != null)
		{
			if (_flags.GetFlagValue(Row, Column, CellFlags.RichText))
			{
				return Cells[Row, Column].RichText.Text;
			}
			return valueInner;
		}
		return null;
	}

	public T GetValue<T>(int Row, int Column)
	{
		CheckSheetType();
		object valueInner = GetValueInner(Row, Column);
		if (valueInner == null)
		{
			return default(T);
		}
		if (_flags.GetFlagValue(Row, Column, CellFlags.RichText))
		{
			return (T)(object)Cells[Row, Column].RichText.Text;
		}
		return ConvertUtil.GetTypedCellValue<T>(valueInner);
	}

	public void SetValue(int Row, int Column, object Value)
	{
		CheckSheetType();
		if (Row < 1 || Column < 1 || (Row > 1048576 && Column > 16384))
		{
			throw new ArgumentOutOfRangeException("Row or Column out of range");
		}
		SetValueInner(Row, Column, Value);
	}

	public void SetValue(string Address, object Value)
	{
		CheckSheetType();
		ExcelCellBase.GetRowCol(Address, out var row, out var col, throwException: true);
		if (row < 1 || col < 1 || (row > 1048576 && col > 16384))
		{
			throw new ArgumentOutOfRangeException("Address is invalid or out of range");
		}
		SetValueInner(row, col, Value);
	}

	public int GetMergeCellId(int row, int column)
	{
		for (int i = 0; i < _mergedCells.Count; i++)
		{
			if (!string.IsNullOrEmpty(_mergedCells[i]))
			{
				ExcelRange excelRange = Cells[_mergedCells[i]];
				if (excelRange.Start.Row <= row && row <= excelRange.End.Row && excelRange.Start.Column <= column && column <= excelRange.End.Column)
				{
					return i + 1;
				}
			}
		}
		return 0;
	}

	private void UpdateCrossSheetReferences(string sheetWhoseReferencesShouldBeUpdated, int rowFrom, int rows, int columnFrom, int columns)
	{
		lock (this)
		{
			foreach (Formulas value in _sharedFormulas.Values)
			{
				value.Formula = ExcelCellBase.UpdateFormulaReferences(value.Formula, rows, columns, rowFrom, columnFrom, Name, sheetWhoseReferencesShouldBeUpdated);
			}
			CellsStoreEnumerator<object> cellsStoreEnumerator = new CellsStoreEnumerator<object>(_formulas);
			while (cellsStoreEnumerator.Next())
			{
				if (cellsStoreEnumerator.Value is string)
				{
					cellsStoreEnumerator.Value = ExcelCellBase.UpdateFormulaReferences(cellsStoreEnumerator.Value.ToString(), rows, columns, rowFrom, columnFrom, Name, sheetWhoseReferencesShouldBeUpdated);
				}
			}
		}
	}

	private void UpdateCrossSheetReferenceNames(string oldName, string newName)
	{
		if (string.IsNullOrEmpty(oldName))
		{
			throw new ArgumentNullException("oldName");
		}
		if (string.IsNullOrEmpty(newName))
		{
			throw new ArgumentNullException("newName");
		}
		lock (this)
		{
			foreach (Formulas value in _sharedFormulas.Values)
			{
				value.Formula = ExcelCellBase.UpdateFormulaSheetReferences(value.Formula, oldName, newName);
			}
			CellsStoreEnumerator<object> cellsStoreEnumerator = new CellsStoreEnumerator<object>(_formulas);
			while (cellsStoreEnumerator.Next())
			{
				if (cellsStoreEnumerator.Value is string)
				{
					cellsStoreEnumerator.Value = ExcelCellBase.UpdateFormulaSheetReferences(cellsStoreEnumerator.Value.ToString(), oldName, newName);
				}
			}
		}
	}

	internal void Save()
	{
		DeletePrinterSettings();
		if (_worksheetXml != null && !(this is ExcelChartsheet))
		{
			if (_headerFooter != null)
			{
				HeaderFooter.Save();
			}
			ExcelAddressBase dimension = Dimension;
			if (dimension == null)
			{
				DeleteAllNode("d:dimension/@ref");
			}
			else
			{
				SetXmlNodeString("d:dimension/@ref", dimension.Address);
			}
			if (Drawings.Count == 0)
			{
				DeleteNode("d:drawing");
			}
			SaveComments();
			HeaderFooter.SaveHeaderFooterImages();
			SaveTables();
			SavePivotTables();
		}
		if (!(Drawings.UriDrawing != null))
		{
			return;
		}
		if (Drawings.Count == 0)
		{
			Part.DeleteRelationship(Drawings._drawingRelation.Id);
			_package.Package.DeletePart(Drawings.UriDrawing);
			return;
		}
		foreach (ExcelDrawing drawing in Drawings)
		{
			drawing.AdjustPositionAndSize();
			if (drawing is ExcelChart)
			{
				ExcelChart excelChart = (ExcelChart)drawing;
				excelChart.ChartXml.Save(excelChart.Part.GetStream(FileMode.Create, FileAccess.Write));
			}
		}
		ZipPackagePart part = Drawings.Part;
		Drawings.DrawingXml.Save(part.GetStream(FileMode.Create, FileAccess.Write));
	}

	internal void SaveHandler(ZipOutputStream stream, CompressionLevel compressionLevel, string fileName)
	{
		stream.CodecBufferSize = 8096;
		stream.CompressionLevel = (OfficeOpenXml.Packaging.Ionic.Zlib.CompressionLevel)compressionLevel;
		stream.PutNextEntry(fileName);
		SaveXml(stream);
	}

	private void DeletePrinterSettings()
	{
		XmlAttribute xmlAttribute = (XmlAttribute)WorksheetXml.SelectSingleNode("//d:pageSetup/@r:id", base.NameSpaceManager);
		if (xmlAttribute == null)
		{
			return;
		}
		string value = xmlAttribute.Value;
		xmlAttribute.OwnerElement.Attributes.Remove(xmlAttribute);
		if (Part.RelationshipExists(value))
		{
			ZipPackageRelationship relationship = Part.GetRelationship(value);
			Uri uri = UriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
			Part.DeleteRelationship(relationship.Id);
			if (_package.Package.PartExists(uri))
			{
				_package.Package.DeletePart(uri);
			}
		}
	}

	private void SaveComments()
	{
		if (_comments != null)
		{
			if (_comments.Count == 0)
			{
				if (_comments.Uri != null)
				{
					Part.DeleteRelationship(_comments.RelId);
					_package.Package.DeletePart(_comments.Uri);
				}
				RemoveLegacyDrawingRel(VmlDrawingsComments.RelId);
			}
			else
			{
				if (_comments.Uri == null)
				{
					_comments.Uri = new Uri($"/xl/comments{SheetID}.xml", UriKind.Relative);
				}
				if (_comments.Part == null)
				{
					_comments.Part = _package.Package.CreatePart(_comments.Uri, "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml", _package.Compression);
					Part.CreateRelationship(UriHelper.GetRelativeUri(WorksheetUri, _comments.Uri), TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments");
				}
				_comments.CommentXml.Save(_comments.Part.GetStream(FileMode.Create));
			}
		}
		if (_vmlDrawings == null)
		{
			return;
		}
		if (_vmlDrawings.Count == 0)
		{
			if (_vmlDrawings.Uri != null)
			{
				Part.DeleteRelationship(_vmlDrawings.RelId);
				_package.Package.DeletePart(_vmlDrawings.Uri);
			}
			return;
		}
		if (_vmlDrawings.Uri == null)
		{
			_vmlDrawings.Uri = XmlHelper.GetNewUri(_package.Package, "/xl/drawings/vmlDrawing{0}.vml");
		}
		if (_vmlDrawings.Part == null)
		{
			_vmlDrawings.Part = _package.Package.CreatePart(_vmlDrawings.Uri, "application/vnd.openxmlformats-officedocument.vmlDrawing", _package.Compression);
			ZipPackageRelationship zipPackageRelationship = Part.CreateRelationship(UriHelper.GetRelativeUri(WorksheetUri, _vmlDrawings.Uri), TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing");
			SetXmlNodeString("d:legacyDrawing/@r:id", zipPackageRelationship.Id);
			_vmlDrawings.RelId = zipPackageRelationship.Id;
		}
		_vmlDrawings.VmlDrawingXml.Save(_vmlDrawings.Part.GetStream(FileMode.Create));
	}

	private void SaveTables()
	{
		foreach (ExcelTable table in Tables)
		{
			if (table.ShowHeader || table.ShowTotal)
			{
				int num = table.Address._fromCol;
				HashSet<string> hashSet = new HashSet<string>();
				foreach (ExcelTableColumn item in (IEnumerable<ExcelTableColumn>)table.Columns)
				{
					string text = item.Name.ToLowerInvariant();
					if (table.ShowHeader)
					{
						text = table.WorkSheet.GetValue<string>(table.Address._fromRow, table.Address._fromCol + item.Position);
						if (string.IsNullOrEmpty(text))
						{
							text = item.Name.ToLowerInvariant();
							SetValueInner(table.Address._fromRow, num, ConvertUtil.ExcelDecodeString(item.Name));
						}
						else
						{
							item.Name = text;
						}
					}
					else
					{
						text = item.Name.ToLowerInvariant();
					}
					if (hashSet.Contains(text))
					{
						throw new InvalidDataException($"Table {table.Name} Column {item.Name} does not have a unique name.");
					}
					hashSet.Add(text);
					if (!string.IsNullOrEmpty(item.CalculatedColumnFormula))
					{
						int num2 = (table.ShowHeader ? (table.Address._fromRow + 1) : table.Address._fromRow);
						int num3 = (table.ShowTotal ? (table.Address._toRow - 1) : table.Address._toRow);
						string text2 = ExcelCellBase.TranslateToR1C1(item.CalculatedColumnFormula, num2, num);
						bool flag = text2 != item.CalculatedColumnFormula;
						for (int i = num2; i <= num3; i++)
						{
							SetFormula(i, num, flag ? ExcelCellBase.TranslateFromR1C1(text2, i, num) : text2);
						}
					}
					num++;
				}
			}
			if (table.Part == null)
			{
				int id = table.Id;
				table.TableUri = XmlHelper.GetNewUri(_package.Package, "/xl/tables/table{0}.xml", ref id);
				table.Id = id;
				table.Part = _package.Package.CreatePart(table.TableUri, "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml", Workbook._package.Compression);
				MemoryStream stream = table.Part.GetStream(FileMode.Create);
				table.TableXml.Save(stream);
				ZipPackageRelationship zipPackageRelationship = Part.CreateRelationship(UriHelper.GetRelativeUri(WorksheetUri, table.TableUri), TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table");
				table.RelationshipID = zipPackageRelationship.Id;
				CreateNode("d:tableParts");
				XmlNode xmlNode = base.TopNode.SelectSingleNode("d:tableParts", base.NameSpaceManager);
				XmlElement xmlElement = xmlNode.OwnerDocument.CreateElement("tablePart", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
				xmlNode.AppendChild(xmlElement);
				xmlElement.SetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", zipPackageRelationship.Id);
			}
			else
			{
				MemoryStream stream2 = table.Part.GetStream(FileMode.Create);
				table.TableXml.Save(stream2);
			}
		}
	}

	internal void SetTableTotalFunction(ExcelTable tbl, ExcelTableColumn col, int colNum = -1)
	{
		if (!tbl.ShowTotal)
		{
			return;
		}
		if (colNum == -1)
		{
			for (int i = 0; i < tbl.Columns.Count; i++)
			{
				if (tbl.Columns[i].Name == col.Name)
				{
					colNum = tbl.Address._fromCol + i;
				}
			}
		}
		if (col.TotalsRowFunction == RowFunctions.Custom)
		{
			SetFormula(tbl.Address._toRow, colNum, col.TotalsRowFormula);
		}
		else if (col.TotalsRowFunction != RowFunctions.None)
		{
			switch (col.TotalsRowFunction)
			{
			case RowFunctions.Average:
				SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "101"));
				break;
			case RowFunctions.CountNums:
				SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "102"));
				break;
			case RowFunctions.Count:
				SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "103"));
				break;
			case RowFunctions.Max:
				SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "104"));
				break;
			case RowFunctions.Min:
				SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "105"));
				break;
			case RowFunctions.StdDev:
				SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "107"));
				break;
			case RowFunctions.Var:
				SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "110"));
				break;
			case RowFunctions.Sum:
				SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "109"));
				break;
			default:
				throw new Exception("Unknown RowFunction enum");
			}
		}
		else
		{
			SetValueInner(tbl.Address._toRow, colNum, col.TotalsRowLabel);
		}
	}

	internal void SetFormula(int row, int col, object value)
	{
		_formulas.SetValue(row, col, value);
		if (!ExistsValueInner(row, col))
		{
			SetValueInner(row, col, null);
		}
	}

	private void SavePivotTables()
	{
		foreach (ExcelPivotTable pivotTable in PivotTables)
		{
			if (pivotTable.DataFields.Count > 1)
			{
				XmlElement xmlElement;
				if (pivotTable.DataOnRows)
				{
					xmlElement = pivotTable.PivotTableXml.SelectSingleNode("//d:rowFields", pivotTable.NameSpaceManager) as XmlElement;
					if (xmlElement == null)
					{
						pivotTable.CreateNode("d:rowFields");
						xmlElement = pivotTable.PivotTableXml.SelectSingleNode("//d:rowFields", pivotTable.NameSpaceManager) as XmlElement;
					}
				}
				else
				{
					xmlElement = pivotTable.PivotTableXml.SelectSingleNode("//d:colFields", pivotTable.NameSpaceManager) as XmlElement;
					if (xmlElement == null)
					{
						pivotTable.CreateNode("d:colFields");
						xmlElement = pivotTable.PivotTableXml.SelectSingleNode("//d:colFields", pivotTable.NameSpaceManager) as XmlElement;
					}
				}
				if (xmlElement.SelectSingleNode("d:field[@ x= \"-2\"]", pivotTable.NameSpaceManager) == null)
				{
					XmlElement xmlElement2 = pivotTable.PivotTableXml.CreateElement("field", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
					xmlElement2.SetAttribute("x", "-2");
					xmlElement.AppendChild(xmlElement2);
				}
			}
			pivotTable.SetXmlNodeString("d:location/@ref", pivotTable.Address.Address);
			if (pivotTable.CacheDefinition.SourceRange != null)
			{
				ExcelTable excelTable = null;
				if (pivotTable.CacheDefinition.SourceRange.IsName)
				{
					pivotTable.CacheDefinition.DeleteNode("d:cacheSource/d:worksheetSource/@ref");
					pivotTable.CacheDefinition.SetXmlNodeString("d:cacheSource/d:worksheetSource/@name", ((ExcelNamedRange)pivotTable.CacheDefinition.SourceRange).Name);
				}
				else
				{
					excelTable = Workbook.Worksheets[pivotTable.CacheDefinition.SourceRange.WorkSheet].Tables.GetFromRange(pivotTable.CacheDefinition.SourceRange);
					if (excelTable == null)
					{
						pivotTable.CacheDefinition.DeleteNode("d:cacheSource/d:worksheetSource/@name");
						pivotTable.CacheDefinition.SetXmlNodeString("d:cacheSource/d:worksheetSource/@ref", pivotTable.CacheDefinition.SourceRange.Address);
					}
					else
					{
						pivotTable.CacheDefinition.DeleteNode("d:cacheSource/d:worksheetSource/@ref");
						pivotTable.CacheDefinition.SetXmlNodeString("d:cacheSource/d:worksheetSource/@name", excelTable.Name);
					}
				}
				XmlNodeList xmlNodeList = pivotTable.CacheDefinition.CacheDefinitionXml.SelectNodes("d:pivotCacheDefinition/d:cacheFields/d:cacheField", base.NameSpaceManager);
				int num = 0;
				if (xmlNodeList != null)
				{
					HashSet<string> hashSet = new HashSet<string>();
					foreach (XmlElement item in xmlNodeList)
					{
						if (num >= pivotTable.CacheDefinition.SourceRange.Columns)
						{
							break;
						}
						string text = item.GetAttribute("name");
						if (string.IsNullOrEmpty(text))
						{
							text = ((excelTable == null) ? pivotTable.CacheDefinition.SourceRange.Offset(0, num++, 1, 1).Value.ToString() : excelTable.Columns[num++].Name);
						}
						if (hashSet.Contains(text))
						{
							text = GetNewName(hashSet, text);
						}
						hashSet.Add(text);
						item.SetAttribute("name", text);
					}
					foreach (ExcelPivotTableDataField dataField in pivotTable.DataFields)
					{
						if (string.IsNullOrEmpty(dataField.Name))
						{
							string text2 = ((dataField.Function != DataFieldFunctions.None) ? (dataField.Function.ToString() + " of " + dataField.Field.Name) : dataField.Field.Name);
							string name = text2;
							int num2 = 2;
							while (pivotTable.DataFields.ExistsDfName(name, dataField))
							{
								name = text2 + num2++.ToString(CultureInfo.InvariantCulture);
							}
							dataField.Name = name;
						}
					}
				}
			}
			pivotTable.PivotTableXml.Save(pivotTable.Part.GetStream(FileMode.Create));
			pivotTable.CacheDefinition.CacheDefinitionXml.Save(pivotTable.CacheDefinition.Part.GetStream(FileMode.Create));
		}
	}

	private string GetNewName(HashSet<string> flds, string fldName)
	{
		int num = 2;
		while (flds.Contains(fldName + num.ToString(CultureInfo.InvariantCulture)))
		{
			num++;
		}
		return fldName + num.ToString(CultureInfo.InvariantCulture);
	}

	private static string GetTotalFunction(ExcelTableColumn col, string FunctionNum)
	{
		string arg = Regex.Replace(col.Name, "[\\[\\]#']", (Match m) => "'" + m.Value);
		return $"SUBTOTAL({FunctionNum},{col._tbl.Name}[{arg}])";
	}

	private void SaveXml(Stream stream)
	{
		StreamWriter streamWriter = new StreamWriter(stream, Encoding.UTF8, 65536);
		if (this is ExcelChartsheet)
		{
			streamWriter.Write(_worksheetXml.OuterXml);
		}
		else
		{
			CreateNode("d:cols");
			CreateNode("d:sheetData");
			CreateNode("d:mergeCells");
			CreateNode("d:hyperlinks");
			CreateNode("d:rowBreaks");
			CreateNode("d:colBreaks");
			string outerXml = _worksheetXml.OuterXml;
			int start = 0;
			int end = 0;
			GetBlockPos(outerXml, "cols", ref start, ref end);
			streamWriter.Write(outerXml.Substring(0, start));
			new List<int>();
			UpdateColumnData(streamWriter);
			int start2 = end;
			int end2 = end;
			GetBlockPos(outerXml, "sheetData", ref start2, ref end2);
			streamWriter.Write(outerXml.Substring(end, start2 - end));
			new List<int>();
			UpdateRowCellData(streamWriter);
			int start3 = end2;
			int end3 = end2;
			GetBlockPos(outerXml, "mergeCells", ref start3, ref end3);
			streamWriter.Write(outerXml.Substring(end2, start3 - end2));
			CleanupMergedCells(_mergedCells);
			if (_mergedCells.Count > 0)
			{
				UpdateMergedCells(streamWriter);
			}
			int start4 = end3;
			int end4 = end3;
			GetBlockPos(outerXml, "hyperlinks", ref start4, ref end4);
			streamWriter.Write(outerXml.Substring(end3, start4 - end3));
			UpdateHyperLinks(streamWriter);
			int start5 = end4;
			int end5 = end4;
			GetBlockPos(outerXml, "rowBreaks", ref start5, ref end5);
			streamWriter.Write(outerXml.Substring(end4, start5 - end4));
			UpdateRowBreaks(streamWriter);
			int start6 = end5;
			int end6 = end5;
			GetBlockPos(outerXml, "colBreaks", ref start6, ref end6);
			streamWriter.Write(outerXml.Substring(end5, start6 - end5));
			UpdateColBreaks(streamWriter);
			streamWriter.Write(outerXml.Substring(end6, outerXml.Length - end6));
		}
		streamWriter.Flush();
	}

	private void CleanupMergedCells(MergeCellsCollection _mergedCells)
	{
		int num = 0;
		while (num < _mergedCells.List.Count)
		{
			if (_mergedCells[num] == null)
			{
				_mergedCells.List.RemoveAt(num);
			}
			else
			{
				num++;
			}
		}
	}

	private void UpdateColBreaks(StreamWriter sw)
	{
		StringBuilder stringBuilder = new StringBuilder();
		int num = 0;
		CellsStoreEnumerator<ExcelCoreValue> cellsStoreEnumerator = new CellsStoreEnumerator<ExcelCoreValue>(_values, 0, 0, 0, 16384);
		while (cellsStoreEnumerator.Next())
		{
			if (cellsStoreEnumerator.Value._value is ExcelColumn { PageBreak: not false })
			{
				stringBuilder.AppendFormat("<brk id=\"{0}\" max=\"16383\" man=\"1\"/>", cellsStoreEnumerator.Column);
				num++;
			}
		}
		if (num > 0)
		{
			sw.Write(string.Format("<colBreaks count=\"{0}\" manualBreakCount=\"{0}\">{1}</colBreaks>", num, stringBuilder.ToString()));
		}
	}

	private void UpdateRowBreaks(StreamWriter sw)
	{
		StringBuilder stringBuilder = new StringBuilder();
		int num = 0;
		CellsStoreEnumerator<ExcelCoreValue> cellsStoreEnumerator = new CellsStoreEnumerator<ExcelCoreValue>(_values, 0, 0, 1048576, 0);
		while (cellsStoreEnumerator.Next())
		{
			if (cellsStoreEnumerator.Value._value is RowInternal { PageBreak: not false })
			{
				stringBuilder.AppendFormat("<brk id=\"{0}\" max=\"1048575\" man=\"1\"/>", cellsStoreEnumerator.Row);
				num++;
			}
		}
		if (num > 0)
		{
			sw.Write(string.Format("<rowBreaks count=\"{0}\" manualBreakCount=\"{0}\">{1}</rowBreaks>", num, stringBuilder.ToString()));
		}
	}

	private void UpdateColumnData(StreamWriter sw)
	{
		CellsStoreEnumerator<ExcelCoreValue> cellsStoreEnumerator = new CellsStoreEnumerator<ExcelCoreValue>(_values, 0, 1, 0, 16384);
		bool flag = true;
		while (cellsStoreEnumerator.Next())
		{
			if (flag)
			{
				sw.Write("<cols>");
				flag = false;
			}
			ExcelColumn excelColumn = cellsStoreEnumerator.Value._value as ExcelColumn;
			ExcelStyleCollection<ExcelXfs> cellXfs = _package.Workbook.Styles.CellXfs;
			sw.Write("<col min=\"{0}\" max=\"{1}\"", excelColumn.ColumnMin, excelColumn.ColumnMax);
			if (excelColumn.Hidden)
			{
				sw.Write(" hidden=\"1\"");
			}
			else if (excelColumn.BestFit)
			{
				sw.Write(" bestFit=\"1\"");
			}
			sw.Write(string.Format(CultureInfo.InvariantCulture, " width=\"{0}\" customWidth=\"1\"", new object[1] { excelColumn.Width }));
			if (excelColumn.OutlineLevel > 0)
			{
				sw.Write(" outlineLevel=\"{0}\" ", excelColumn.OutlineLevel);
				if (excelColumn.Collapsed)
				{
					if (excelColumn.Hidden)
					{
						sw.Write(" collapsed=\"1\"");
					}
					else
					{
						sw.Write(" collapsed=\"1\" hidden=\"1\"");
					}
				}
			}
			if (excelColumn.Phonetic)
			{
				sw.Write(" phonetic=\"1\"");
			}
			int num = ((excelColumn.StyleID >= 0) ? cellXfs[excelColumn.StyleID].newID : excelColumn.StyleID);
			if (num > 0)
			{
				sw.Write(" style=\"{0}\"", num);
			}
			sw.Write("/>");
		}
		if (!flag)
		{
			sw.Write("</cols>");
		}
	}

	private void UpdateRowCellData(StreamWriter sw)
	{
		ExcelStyleCollection<ExcelXfs> cellXfs = _package.Workbook.Styles.CellXfs;
		int num = -1;
		new StringBuilder();
		Dictionary<string, ExcelWorkbook.SharedStringItem> sharedStrings = _package.Workbook._sharedStrings;
		_ = _package.Workbook.Styles;
		StringBuilder stringBuilder = new StringBuilder();
		stringBuilder.Append("<sheetData>");
		FixSharedFormulas();
		columnStyles = new Dictionary<int, int>();
		CellsStoreEnumerator<ExcelCoreValue> cellsStoreEnumerator = new CellsStoreEnumerator<ExcelCoreValue>(_values, 1, 0, 1048576, 16384);
		while (cellsStoreEnumerator.Next())
		{
			if (cellsStoreEnumerator.Column > 0)
			{
				ExcelCoreValue value = cellsStoreEnumerator.Value;
				int newID = cellXfs[(value._styleId == 0) ? GetStyleIdDefaultWithMemo(cellsStoreEnumerator.Row, cellsStoreEnumerator.Column) : value._styleId].newID;
				if (cellsStoreEnumerator.Row != num)
				{
					WriteRow(stringBuilder, cellXfs, num, cellsStoreEnumerator.Row);
					num = cellsStoreEnumerator.Row;
				}
				object obj = value._value;
				object value2 = _formulas.GetValue(cellsStoreEnumerator.Row, cellsStoreEnumerator.Column);
				if (value2 is int num2)
				{
					Formulas formulas = _sharedFormulas[num2];
					if (formulas.Address.IndexOf(':') > 0)
					{
						if (formulas.StartCol == cellsStoreEnumerator.Column && formulas.StartRow == cellsStoreEnumerator.Row)
						{
							if (formulas.IsArray)
							{
								stringBuilder.AppendFormat("<c r=\"{0}\" s=\"{1}\"{5}><f ref=\"{2}\" t=\"array\">{3}</f>{4}</c>", cellsStoreEnumerator.CellAddress, (newID >= 0) ? newID : 0, formulas.Address, ConvertUtil.ExcelEscapeString(formulas.Formula), GetFormulaValue(obj), GetCellType(obj, allowStr: true));
							}
							else
							{
								stringBuilder.AppendFormat("<c r=\"{0}\" s=\"{1}\"{6}><f ref=\"{2}\" t=\"shared\" si=\"{3}\">{4}</f>{5}</c>", cellsStoreEnumerator.CellAddress, (newID >= 0) ? newID : 0, formulas.Address, num2, ConvertUtil.ExcelEscapeString(formulas.Formula), GetFormulaValue(obj), GetCellType(obj, allowStr: true));
							}
						}
						else if (formulas.IsArray)
						{
							stringBuilder.AppendFormat("<c r=\"{0}\" s=\"{1}\"/>", cellsStoreEnumerator.CellAddress, (newID >= 0) ? newID : 0);
						}
						else
						{
							stringBuilder.AppendFormat("<c r=\"{0}\" s=\"{1}\"{4}><f t=\"shared\" si=\"{2}\"/>{3}</c>", cellsStoreEnumerator.CellAddress, (newID >= 0) ? newID : 0, num2, GetFormulaValue(obj), GetCellType(obj, allowStr: true));
						}
					}
					else if (formulas.IsArray)
					{
						stringBuilder.AppendFormat("<c r=\"{0}\" s=\"{1}\"{5}><f ref=\"{2}\" t=\"array\">{3}</f>{4}</c>", cellsStoreEnumerator.CellAddress, (newID >= 0) ? newID : 0, $"{formulas.Address}:{formulas.Address}", ConvertUtil.ExcelEscapeString(formulas.Formula), GetFormulaValue(obj), GetCellType(obj, allowStr: true));
					}
					else
					{
						stringBuilder.AppendFormat("<c r=\"{0}\" s=\"{1}\"{2}>", formulas.Address, (newID >= 0) ? newID : 0, GetCellType(obj, allowStr: true));
						stringBuilder.AppendFormat("<f>{0}</f>{1}</c>", ConvertUtil.ExcelEscapeString(formulas.Formula), GetFormulaValue(obj));
					}
				}
				else if (value2 != null && value2.ToString() != "")
				{
					stringBuilder.AppendFormat("<c r=\"{0}\" s=\"{1}\"{2}>", cellsStoreEnumerator.CellAddress, (newID >= 0) ? newID : 0, GetCellType(obj, allowStr: true));
					stringBuilder.AppendFormat("<f>{0}</f>{1}</c>", ConvertUtil.ExcelEscapeString(value2.ToString()), GetFormulaValue(obj));
				}
				else if (obj == null && newID > 0)
				{
					stringBuilder.AppendFormat("<c r=\"{0}\" s=\"{1}\"/>", cellsStoreEnumerator.CellAddress, (newID >= 0) ? newID : 0);
				}
				else if (obj != null)
				{
					if (obj is IEnumerable enumerable && !(obj is string))
					{
						IEnumerator enumerator = enumerable.GetEnumerator();
						obj = ((!enumerator.MoveNext() || enumerator.Current == null) ? string.Empty : enumerator.Current);
					}
					if (TypeCompat.IsPrimitive(obj) || obj is double || obj is decimal || obj is DateTime || obj is TimeSpan)
					{
						stringBuilder.AppendFormat("<c r=\"{0}\" s=\"{1}\"{2}>", cellsStoreEnumerator.CellAddress, (newID >= 0) ? newID : 0, GetCellType(obj));
						stringBuilder.AppendFormat("{0}</c>", GetFormulaValue(obj));
					}
					else
					{
						string key = Convert.ToString(obj);
						int num3;
						if (!sharedStrings.ContainsKey(key))
						{
							num3 = sharedStrings.Count;
							sharedStrings.Add(key, new ExcelWorkbook.SharedStringItem
							{
								isRichText = _flags.GetFlagValue(cellsStoreEnumerator.Row, cellsStoreEnumerator.Column, CellFlags.RichText),
								pos = num3
							});
						}
						else
						{
							num3 = sharedStrings[key].pos;
						}
						stringBuilder.AppendFormat("<c r=\"{0}\" s=\"{1}\" t=\"s\">", cellsStoreEnumerator.CellAddress, (newID >= 0) ? newID : 0);
						stringBuilder.AppendFormat("<v>{0}</v></c>", num3);
					}
				}
			}
			else
			{
				WriteRow(stringBuilder, cellXfs, num, cellsStoreEnumerator.Row);
				num = cellsStoreEnumerator.Row;
			}
			if (stringBuilder.Length > 6291456)
			{
				sw.Write(stringBuilder.ToString());
				sw.Flush();
				stringBuilder.Length = 0;
			}
		}
		columnStyles = null;
		if (num != -1)
		{
			stringBuilder.Append("</row>");
		}
		stringBuilder.Append("</sheetData>");
		sw.Write(stringBuilder.ToString());
		sw.Flush();
	}

	private void FixSharedFormulas()
	{
		List<int> list = new List<int>();
		foreach (Formulas value3 in _sharedFormulas.Values)
		{
			ExcelAddressBase excelAddressBase = new ExcelAddressBase(value3.Address);
			object value = _formulas.GetValue(excelAddressBase._fromRow, excelAddressBase._fromCol);
			if (value is int && (!(value is int) || (int)value == value3.Index))
			{
				continue;
			}
			for (int j = excelAddressBase._fromRow; j <= excelAddressBase._toRow; j++)
			{
				for (int k = excelAddressBase._fromCol; k <= excelAddressBase._toCol; k++)
				{
					if (excelAddressBase._fromRow != j || excelAddressBase._fromCol != k)
					{
						object value2 = _formulas.GetValue(j, k);
						if (value2 is int && (int)value2 == value3.Index)
						{
							_formulas.SetValue(j, k, value3.GetFormula(j, k, Name));
						}
					}
				}
			}
			list.Add(value3.Index);
		}
		list.ForEach(delegate(int i)
		{
			_sharedFormulas.Remove(i);
		});
	}

	internal int GetStyleIdDefaultWithMemo(int row, int col)
	{
		int styleId = 0;
		if (ExistsStyleInner(row, 0, ref styleId))
		{
			return styleId;
		}
		if (!columnStyles.ContainsKey(col))
		{
			if (ExistsStyleInner(0, col, ref styleId))
			{
				columnStyles.Add(col, styleId);
			}
			else
			{
				int row2 = 0;
				int col2 = col;
				if (_values.PrevCell(ref row2, ref col2))
				{
					ExcelCoreValue value = _values.GetValue(0, col2);
					ExcelColumn excelColumn = (ExcelColumn)value._value;
					if (excelColumn != null && excelColumn.ColumnMax >= col)
					{
						columnStyles.Add(col, value._styleId);
					}
					else
					{
						columnStyles.Add(col, 0);
					}
				}
				else
				{
					columnStyles.Add(col, 0);
				}
			}
		}
		return columnStyles[col];
	}

	private object GetFormulaValue(object v)
	{
		if (v != null && v.ToString() != "")
		{
			return "<v>" + ConvertUtil.ExcelEscapeString(GetValueForXml(v)) + "</v>";
		}
		return "";
	}

	private string GetCellType(object v, bool allowStr = false)
	{
		if (v is bool)
		{
			return " t=\"b\"";
		}
		if ((v is double && double.IsInfinity((double)v)) || v is ExcelErrorValue)
		{
			return " t=\"e\"";
		}
		if (allowStr && v != null && !TypeCompat.IsPrimitive(v) && !(v is double) && !(v is decimal) && !(v is DateTime) && !(v is TimeSpan))
		{
			return " t=\"str\"";
		}
		return "";
	}

	private string GetValueForXml(object v)
	{
		try
		{
			if (v is DateTime dateTime)
			{
				double num = dateTime.ToOADate();
				if (Workbook.Date1904)
				{
					num -= 1462.0;
				}
				return num.ToString(CultureInfo.InvariantCulture);
			}
			if (v is TimeSpan)
			{
				return DateTime.FromOADate(0.0).Add((TimeSpan)v).ToOADate()
					.ToString(CultureInfo.InvariantCulture);
			}
			if (TypeCompat.IsPrimitive(v) || v is double || v is decimal)
			{
				if (v is double && double.IsNaN((double)v))
				{
					return "";
				}
				if (v is double && double.IsInfinity((double)v))
				{
					return "#NUM!";
				}
				return Convert.ToDouble(v, CultureInfo.InvariantCulture).ToString("R15", CultureInfo.InvariantCulture);
			}
			return v.ToString();
		}
		catch
		{
			return "0";
		}
	}

	private void WriteRow(StringBuilder cache, ExcelStyleCollection<ExcelXfs> cellXfs, int prevRow, int row)
	{
		if (prevRow != -1)
		{
			cache.Append("</row>");
		}
		cache.AppendFormat("<row r=\"{0}\"", row);
		if (GetValueInner(row, 0) is RowInternal rowInternal)
		{
			if (rowInternal.Hidden)
			{
				cache.Append(" hidden=\"1\"");
			}
			if (rowInternal.Height >= 0.0)
			{
				cache.AppendFormat(string.Format(CultureInfo.InvariantCulture, " ht=\"{0}\"", new object[1] { rowInternal.Height }));
				if (rowInternal.CustomHeight)
				{
					cache.Append(" customHeight=\"1\"");
				}
			}
			if (rowInternal.OutlineLevel > 0)
			{
				cache.AppendFormat(" outlineLevel =\"{0}\"", rowInternal.OutlineLevel);
				if (rowInternal.Collapsed)
				{
					if (rowInternal.Hidden)
					{
						cache.Append(" collapsed=\"1\"");
					}
					else
					{
						cache.Append(" collapsed=\"1\" hidden=\"1\"");
					}
				}
			}
			if (rowInternal.Phonetic)
			{
				cache.Append(" ph=\"1\"");
			}
		}
		int styleInner = GetStyleInner(row, 0);
		if (styleInner > 0)
		{
			cache.AppendFormat(" s=\"{0}\" customFormat=\"1\"", cellXfs[styleInner].newID);
		}
		cache.Append(">");
	}

	private void WriteRow(StreamWriter sw, ExcelStyleCollection<ExcelXfs> cellXfs, int prevRow, int row)
	{
		if (prevRow != -1)
		{
			sw.Write("</row>");
		}
		sw.Write("<row r=\"{0}\"", row);
		if (GetValueInner(row, 0) is RowInternal rowInternal)
		{
			if (rowInternal.Hidden)
			{
				sw.Write(" hidden=\"1\"");
			}
			if (rowInternal.Height >= 0.0)
			{
				sw.Write(string.Format(CultureInfo.InvariantCulture, " ht=\"{0}\"", new object[1] { rowInternal.Height }));
				if (rowInternal.CustomHeight)
				{
					sw.Write(" customHeight=\"1\"");
				}
			}
			if (rowInternal.OutlineLevel > 0)
			{
				sw.Write(" outlineLevel =\"{0}\"", rowInternal.OutlineLevel);
				if (rowInternal.Collapsed)
				{
					if (rowInternal.Hidden)
					{
						sw.Write(" collapsed=\"1\"");
					}
					else
					{
						sw.Write(" collapsed=\"1\" hidden=\"1\"");
					}
				}
			}
			if (rowInternal.Phonetic)
			{
				sw.Write(" ph=\"1\"");
			}
		}
		int styleInner = GetStyleInner(row, 0);
		if (styleInner > 0)
		{
			sw.Write(" s=\"{0}\" customFormat=\"1\"", cellXfs[styleInner].newID);
		}
		sw.Write(">");
	}

	private void UpdateHyperLinks(StreamWriter sw)
	{
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		CellsStoreEnumerator<Uri> cellsStoreEnumerator = new CellsStoreEnumerator<Uri>(_hyperLinks);
		bool flag = true;
		while (cellsStoreEnumerator.Next())
		{
			Uri value = _hyperLinks.GetValue(cellsStoreEnumerator.Row, cellsStoreEnumerator.Column);
			if (flag && value != null)
			{
				sw.Write("<hyperlinks>");
				flag = false;
			}
			if (value is ExcelHyperLink && !string.IsNullOrEmpty((value as ExcelHyperLink).ReferenceAddress))
			{
				ExcelHyperLink excelHyperLink = value as ExcelHyperLink;
				sw.Write("<hyperlink ref=\"{0}\" location=\"{1}\"{2}{3}/>", Cells[cellsStoreEnumerator.Row, cellsStoreEnumerator.Column, cellsStoreEnumerator.Row + excelHyperLink.RowSpann, cellsStoreEnumerator.Column + excelHyperLink.ColSpann].Address, ExcelCellBase.GetFullAddress(SecurityElement.Escape(Name), SecurityElement.Escape(excelHyperLink.ReferenceAddress)), string.IsNullOrEmpty(excelHyperLink.Display) ? "" : (" display=\"" + SecurityElement.Escape(excelHyperLink.Display) + "\""), string.IsNullOrEmpty(excelHyperLink.ToolTip) ? "" : (" tooltip=\"" + SecurityElement.Escape(excelHyperLink.ToolTip) + "\""));
			}
			else
			{
				if (!(value != null))
				{
					continue;
				}
				Uri uri = ((!(value is ExcelHyperLink)) ? value : ((ExcelHyperLink)value).OriginalUri);
				if (dictionary.ContainsKey(uri.OriginalString))
				{
					_ = dictionary[uri.OriginalString];
					continue;
				}
				ZipPackageRelationship zipPackageRelationship = Part.CreateRelationship(uri, TargetMode.External, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink");
				if (value is ExcelHyperLink)
				{
					ExcelHyperLink excelHyperLink2 = value as ExcelHyperLink;
					sw.Write("<hyperlink ref=\"{0}\"{2}{3} r:id=\"{1}\"/>", ExcelCellBase.GetAddress(cellsStoreEnumerator.Row, cellsStoreEnumerator.Column), zipPackageRelationship.Id, string.IsNullOrEmpty(excelHyperLink2.Display) ? "" : (" display=\"" + SecurityElement.Escape(excelHyperLink2.Display) + "\""), string.IsNullOrEmpty(excelHyperLink2.ToolTip) ? "" : (" tooltip=\"" + SecurityElement.Escape(excelHyperLink2.ToolTip) + "\""));
				}
				else
				{
					sw.Write("<hyperlink ref=\"{0}\" r:id=\"{1}\"/>", ExcelCellBase.GetAddress(cellsStoreEnumerator.Row, cellsStoreEnumerator.Column), zipPackageRelationship.Id);
				}
				_ = zipPackageRelationship.Id;
			}
		}
		if (!flag)
		{
			sw.Write("</hyperlinks>");
		}
	}

	private XmlNode CreateHyperLinkCollection()
	{
		XmlElement newChild = _worksheetXml.CreateElement("hyperlinks", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
		XmlNode xmlNode = _worksheetXml.SelectSingleNode("//d:conditionalFormatting", base.NameSpaceManager);
		if (xmlNode == null)
		{
			xmlNode = _worksheetXml.SelectSingleNode("//d:mergeCells", base.NameSpaceManager);
			if (xmlNode == null)
			{
				xmlNode = _worksheetXml.SelectSingleNode("//d:sheetData", base.NameSpaceManager);
			}
		}
		return _worksheetXml.DocumentElement.InsertAfter(newChild, xmlNode);
	}

	internal void ClearValidations()
	{
		_dataValidation = null;
	}

	internal int GetStyleID(string StyleName)
	{
		ExcelNamedStyleXml obj = null;
		Workbook.Styles.NamedStyles.FindByID(StyleName, ref obj);
		if (obj.XfId == int.MinValue)
		{
			obj.XfId = Workbook.Styles.CellXfs.FindIndexByID(obj.Style.Id);
		}
		return obj.XfId;
	}

	internal int GetMaxShareFunctionIndex(bool isArray)
	{
		int i = _sharedFormulas.Count + 1;
		if (isArray)
		{
			i |= 0x40000000;
		}
		for (; _sharedFormulas.ContainsKey(i); i++)
		{
		}
		return i;
	}

	internal void SetHFLegacyDrawingRel(string relID)
	{
		SetXmlNodeString("d:legacyDrawingHF/@r:id", relID);
	}

	internal void RemoveLegacyDrawingRel(string relID)
	{
		XmlNode xmlNode = WorksheetXml.DocumentElement.SelectSingleNode($"d:legacyDrawing[@r:id=\"{relID}\"]", base.NameSpaceManager);
		xmlNode?.ParentNode.RemoveChild(xmlNode);
	}

	internal void UpdateCellsWithDate1904Setting()
	{
		CellsStoreEnumerator<ExcelCoreValue> cellsStoreEnumerator = new CellsStoreEnumerator<ExcelCoreValue>(_values);
		double num = (Workbook.Date1904 ? (-1462.0) : 1462.0);
		while (cellsStoreEnumerator.MoveNext())
		{
			if (cellsStoreEnumerator.Value._value is DateTime)
			{
				try
				{
					double num2 = ((DateTime)cellsStoreEnumerator.Value._value).ToOADate();
					num2 += num;
					SetValueInner(cellsStoreEnumerator.Row, cellsStoreEnumerator.Column, DateTime.FromOADate(num2));
				}
				catch
				{
				}
			}
		}
	}

	internal string GetFormula(int row, int col)
	{
		object value = _formulas.GetValue(row, col);
		if (value is int)
		{
			return _sharedFormulas[(int)value].GetFormula(row, col, Name);
		}
		if (value != null)
		{
			return value.ToString();
		}
		return "";
	}

	internal string GetFormulaR1C1(int row, int col)
	{
		object value = _formulas.GetValue(row, col);
		if (value is int)
		{
			Formulas formulas = _sharedFormulas[(int)value];
			return ExcelCellBase.TranslateToR1C1(formulas.Formula, formulas.StartRow, formulas.StartCol);
		}
		if (value != null)
		{
			return ExcelCellBase.TranslateToR1C1(value.ToString(), row, col);
		}
		return "";
	}

	private void DisposeInternal(IDisposable candidateDisposable)
	{
		candidateDisposable?.Dispose();
	}

	public void Dispose()
	{
		DisposeInternal(_values);
		DisposeInternal(_formulas);
		DisposeInternal(_flags);
		DisposeInternal(_hyperLinks);
		DisposeInternal(_commentsStore);
		DisposeInternal(_formulaTokens);
		_values = null;
		_formulas = null;
		_flags = null;
		_hyperLinks = null;
		_commentsStore = null;
		_formulaTokens = null;
		_package = null;
		_pivotTables = null;
		_protection = null;
		if (_sharedFormulas != null)
		{
			_sharedFormulas.Clear();
		}
		_sharedFormulas = null;
		_sheetView = null;
		_tables = null;
		_vmlDrawings = null;
		_conditionalFormatting = null;
		_dataValidation = null;
		_drawings = null;
	}

	internal ExcelColumn GetColumn(int column)
	{
		ExcelColumn excelColumn = GetValueInner(0, column) as ExcelColumn;
		if (excelColumn == null)
		{
			int row = 0;
			int col = column;
			if (_values.PrevCell(ref row, ref col))
			{
				if (GetValueInner(0, col) is ExcelColumn excelColumn2 && excelColumn2.ColumnMax >= column)
				{
					return excelColumn2;
				}
				return null;
			}
		}
		return excelColumn;
	}

	public bool Equals(ExcelWorksheet x, ExcelWorksheet y)
	{
		if (x.Name == y.Name && x.SheetID == y.SheetID)
		{
			return x.WorksheetXml.OuterXml == y.WorksheetXml.OuterXml;
		}
		return false;
	}

	public int GetHashCode(ExcelWorksheet obj)
	{
		return obj.WorksheetXml.OuterXml.GetHashCode();
	}

	internal object GetValueInner(int row, int col)
	{
		return _values.GetValue(row, col)._value;
	}

	internal int GetStyleInner(int row, int col)
	{
		return _values.GetValue(row, col)._styleId;
	}

	internal void SetValueInner(int row, int col, object value)
	{
		_values.SetValueSpecial(row, col, _setValueInnerUpdateDelegate, value);
	}

	private static void SetValueInnerUpdate(List<ExcelCoreValue> list, int index, object value)
	{
		list[index] = new ExcelCoreValue
		{
			_value = value,
			_styleId = list[index]._styleId
		};
	}

	internal void SetStyleInner(int row, int col, int styleId)
	{
		_values.SetValueSpecial(row, col, SetStyleInnerUpdate, styleId);
	}

	private void SetStyleInnerUpdate(List<ExcelCoreValue> list, int index, object styleId)
	{
		list[index] = new ExcelCoreValue
		{
			_value = list[index]._value,
			_styleId = (int)styleId
		};
	}

	internal void SetRangeValueInner(int fromRow, int fromColumn, int toRow, int toColumn, object[,] values)
	{
		int rowBound = values.GetUpperBound(0);
		int colBound = values.GetUpperBound(1);
		_values.SetRangeValueSpecial(fromRow, fromColumn, toRow, toColumn, delegate(List<ExcelCoreValue> list, int index, int row, int column, object value)
		{
			object value2 = null;
			if (rowBound >= row - fromRow && colBound >= column - fromColumn)
			{
				value2 = values[row - fromRow, column - fromColumn];
			}
			list[index] = new ExcelCoreValue
			{
				_value = value2,
				_styleId = list[index]._styleId
			};
		}, values);
	}

	private void SetRangeValueUpdate(List<ExcelCoreValue> list, int index, int row, int column, object values)
	{
		list[index] = new ExcelCoreValue
		{
			_value = ((object[,])values)[row, column],
			_styleId = list[index]._styleId
		};
	}

	internal bool ExistsValueInner(int row, int col)
	{
		return _values.GetValue(row, col)._value != null;
	}

	internal bool ExistsStyleInner(int row, int col)
	{
		return _values.GetValue(row, col)._styleId != 0;
	}

	internal bool ExistsValueInner(int row, int col, ref object value)
	{
		value = _values.GetValue(row, col)._value;
		return value != null;
	}

	internal bool ExistsStyleInner(int row, int col, ref int styleId)
	{
		styleId = _values.GetValue(row, col)._styleId;
		return styleId != 0;
	}
}
