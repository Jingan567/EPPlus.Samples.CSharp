using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing;

public class EpplusExcelDataProvider : ExcelDataProvider
{
	public class RangeInfo : IRangeInfo, IEnumerator<ICellInfo>, IDisposable, IEnumerator, IEnumerable<ICellInfo>, IEnumerable
	{
		internal ExcelWorksheet _ws;

		private CellsStoreEnumerator<ExcelCoreValue> _values;

		private int _fromRow;

		private int _toRow;

		private int _fromCol;

		private int _toCol;

		private int _cellCount;

		private ExcelAddressBase _address;

		private ICellInfo _cell;

		public bool IsEmpty
		{
			get
			{
				if (_cellCount > 0)
				{
					return false;
				}
				if (_values.Next())
				{
					_values.Reset();
					return false;
				}
				return true;
			}
		}

		public bool IsMulti
		{
			get
			{
				if (_cellCount == 0)
				{
					if (_values.Next() && _values.Next())
					{
						_values.Reset();
						return true;
					}
					_values.Reset();
					return false;
				}
				if (_cellCount > 1)
				{
					return true;
				}
				return false;
			}
		}

		public ICellInfo Current => _cell;

		public ExcelWorksheet Worksheet => _ws;

		object IEnumerator.Current => this;

		public ExcelAddressBase Address => _address;

		public RangeInfo(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
		{
			_ws = ws;
			_fromRow = fromRow;
			_fromCol = fromCol;
			_toRow = toRow;
			_toCol = toCol;
			_address = new ExcelAddressBase(_fromRow, _fromCol, _toRow, _toCol);
			_address._ws = ws.Name;
			_values = new CellsStoreEnumerator<ExcelCoreValue>(ws._values, _fromRow, _fromCol, _toRow, _toCol);
			_cell = new CellInfo(_ws, _values);
		}

		public RangeInfo(ExcelWorksheet ws, ExcelAddressBase address)
		{
			_ws = ws;
			_fromRow = address._fromRow;
			_fromCol = address._fromCol;
			_toRow = address._toRow;
			_toCol = address._toCol;
			_address = address;
			_address._ws = ws.Name;
			_values = new CellsStoreEnumerator<ExcelCoreValue>(ws._values, _fromRow, _fromCol, _toRow, _toCol);
			_cell = new CellInfo(_ws, _values);
		}

		public int GetNCells()
		{
			return (_toRow - _fromRow + 1) * (_toCol - _fromCol + 1);
		}

		public void Dispose()
		{
		}

		public bool MoveNext()
		{
			_cellCount++;
			return _values.MoveNext();
		}

		public void Reset()
		{
			_values.Init();
		}

		public bool NextCell()
		{
			_cellCount++;
			return _values.MoveNext();
		}

		public IEnumerator<ICellInfo> GetEnumerator()
		{
			Reset();
			return this;
		}

		IEnumerator IEnumerable.GetEnumerator()
		{
			return this;
		}

		public object GetValue(int row, int col)
		{
			return _ws.GetValue(row, col);
		}

		public object GetOffset(int rowOffset, int colOffset)
		{
			if (_values.Row < _fromRow || _values.Column < _fromCol)
			{
				return _ws.GetValue(_fromRow + rowOffset, _fromCol + colOffset);
			}
			return _ws.GetValue(_values.Row + rowOffset, _values.Column + colOffset);
		}
	}

	public class CellInfo : ICellInfo
	{
		private ExcelWorksheet _ws;

		private CellsStoreEnumerator<ExcelCoreValue> _values;

		public string Address => _values.CellAddress;

		public int Row => _values.Row;

		public int Column => _values.Column;

		public string Formula => _ws.GetFormula(_values.Row, _values.Column);

		public object Value => _values.Value._value;

		public double ValueDouble => ConvertUtil.GetValueDouble(_values.Value._value, ignoreBool: true);

		public double ValueDoubleLogical => ConvertUtil.GetValueDouble(_values.Value._value);

		public bool IsHiddenRow
		{
			get
			{
				if (_ws.GetValueInner(_values.Row, 0) is RowInternal rowInternal)
				{
					if (!rowInternal.Hidden)
					{
						return rowInternal.Height == 0.0;
					}
					return true;
				}
				return false;
			}
		}

		public bool IsExcelError => ExcelErrorValue.Values.IsErrorValue(_values.Value._value);

		public IList<Token> Tokens => _ws._formulaTokens.GetValue(_values.Row, _values.Column);

		internal CellInfo(ExcelWorksheet ws, CellsStoreEnumerator<ExcelCoreValue> values)
		{
			_ws = ws;
			_values = values;
		}
	}

	public class NameInfo : INameInfo
	{
		public ulong Id { get; set; }

		public string Worksheet { get; set; }

		public string Name { get; set; }

		public string Formula { get; set; }

		public IList<Token> Tokens { get; internal set; }

		public object Value { get; set; }
	}

	private readonly ExcelPackage _package;

	private ExcelWorksheet _currentWorksheet;

	private RangeAddressFactory _rangeAddressFactory;

	private Dictionary<ulong, INameInfo> _names = new Dictionary<ulong, INameInfo>();

	public override int ExcelMaxColumns => 16384;

	public override int ExcelMaxRows => 1048576;

	public EpplusExcelDataProvider(ExcelPackage package)
	{
		_package = package;
		_rangeAddressFactory = new RangeAddressFactory(this);
	}

	public override ExcelNamedRangeCollection GetWorksheetNames(string worksheet)
	{
		return _package.Workbook.Worksheets[worksheet]?.Names;
	}

	public override ExcelNamedRangeCollection GetWorkbookNameValues()
	{
		return _package.Workbook.Names;
	}

	public override IRangeInfo GetRange(string worksheet, int fromRow, int fromCol, int toRow, int toCol)
	{
		SetCurrentWorksheet(worksheet);
		string name = (string.IsNullOrEmpty(worksheet) ? _currentWorksheet.Name : worksheet);
		return new RangeInfo(_package.Workbook.Worksheets[name], fromRow, fromCol, toRow, toCol);
	}

	public override IRangeInfo GetRange(string worksheet, int row, int column, string address)
	{
		ExcelAddress excelAddress = new ExcelAddress(worksheet, address);
		if (excelAddress.Table != null)
		{
			excelAddress = ConvertToA1C1(excelAddress);
		}
		string name = (string.IsNullOrEmpty(excelAddress.WorkSheet) ? _currentWorksheet.Name : excelAddress.WorkSheet);
		return new RangeInfo(_package.Workbook.Worksheets[name], excelAddress);
	}

	public override IRangeInfo GetRange(string worksheet, string address)
	{
		ExcelAddress excelAddress = new ExcelAddress(worksheet, address);
		if (excelAddress.Table != null)
		{
			excelAddress = ConvertToA1C1(excelAddress);
		}
		string name = (string.IsNullOrEmpty(excelAddress.WorkSheet) ? _currentWorksheet.Name : excelAddress.WorkSheet);
		return new RangeInfo(_package.Workbook.Worksheets[name], excelAddress);
	}

	private ExcelAddress ConvertToA1C1(ExcelAddress addr)
	{
		addr.SetRCFromTable(_package, addr);
		return new ExcelAddress(addr._fromRow, addr._fromCol, addr._toRow, addr._toCol)
		{
			_ws = addr._ws
		};
	}

	public override INameInfo GetName(string worksheet, string name)
	{
		ExcelNamedRange excelNamedRange;
		ExcelWorksheet excelWorksheet;
		if (string.IsNullOrEmpty(worksheet))
		{
			if (!_package._workbook.Names.ContainsKey(name))
			{
				return null;
			}
			excelNamedRange = _package._workbook.Names[name];
			excelWorksheet = null;
		}
		else
		{
			excelWorksheet = _package._workbook.Worksheets[worksheet];
			if (excelWorksheet != null && excelWorksheet.Names.ContainsKey(name))
			{
				excelNamedRange = excelWorksheet.Names[name];
			}
			else
			{
				if (!_package._workbook.Names.ContainsKey(name))
				{
					return null;
				}
				excelNamedRange = _package._workbook.Names[name];
			}
		}
		ulong cellID = ExcelCellBase.GetCellID(excelNamedRange.LocalSheetId, excelNamedRange.Index, 0);
		if (_names.ContainsKey(cellID))
		{
			return _names[cellID];
		}
		NameInfo nameInfo = new NameInfo
		{
			Id = cellID,
			Name = name,
			Worksheet = ((excelNamedRange.Worksheet == null) ? excelNamedRange._ws : excelNamedRange.Worksheet.Name),
			Formula = excelNamedRange.Formula
		};
		if (excelNamedRange._fromRow > 0)
		{
			nameInfo.Value = new RangeInfo(excelNamedRange.Worksheet ?? excelWorksheet, excelNamedRange._fromRow, excelNamedRange._fromCol, excelNamedRange._toRow, excelNamedRange._toCol);
		}
		else
		{
			nameInfo.Value = excelNamedRange.Value;
		}
		_names.Add(cellID, nameInfo);
		return nameInfo;
	}

	public override IEnumerable<object> GetRangeValues(string address)
	{
		SetCurrentWorksheet(ExcelAddressInfo.Parse(address));
		ExcelAddress excelAddress = new ExcelAddress(address);
		string name = (string.IsNullOrEmpty(excelAddress.WorkSheet) ? _currentWorksheet.Name : excelAddress.WorkSheet);
		return (IEnumerable<object>)new CellsStoreEnumerator<ExcelCoreValue>(_package.Workbook.Worksheets[name]._values, excelAddress._fromRow, excelAddress._fromCol, excelAddress._toRow, excelAddress._toCol);
	}

	public object GetValue(int row, int column)
	{
		return _currentWorksheet.GetValueInner(row, column);
	}

	public bool IsMerged(int row, int column)
	{
		return _currentWorksheet.MergedCells[row, column] != null;
	}

	public bool IsHidden(int row, int column)
	{
		if (!_currentWorksheet.Column(column).Hidden && _currentWorksheet.Column(column).Width != 0.0 && !_currentWorksheet.Row(row).Hidden)
		{
			return _currentWorksheet.Row(column).Height == 0.0;
		}
		return true;
	}

	public override object GetCellValue(string sheetName, int row, int col)
	{
		SetCurrentWorksheet(sheetName);
		return _currentWorksheet.GetValueInner(row, col);
	}

	public override ExcelCellAddress GetDimensionEnd(string worksheet)
	{
		ExcelCellAddress result = null;
		try
		{
			result = _package.Workbook.Worksheets[worksheet].Dimension.End;
		}
		catch
		{
		}
		return result;
	}

	private void SetCurrentWorksheet(ExcelAddressInfo addressInfo)
	{
		if (addressInfo.WorksheetIsSpecified)
		{
			_currentWorksheet = _package.Workbook.Worksheets[addressInfo.Worksheet];
		}
		else if (_currentWorksheet == null)
		{
			_currentWorksheet = _package.Workbook.Worksheets.First();
		}
	}

	private void SetCurrentWorksheet(string worksheetName)
	{
		if (!string.IsNullOrEmpty(worksheetName))
		{
			_currentWorksheet = _package.Workbook.Worksheets[worksheetName];
		}
		else
		{
			_currentWorksheet = _package.Workbook.Worksheets.First();
		}
	}

	public override void Dispose()
	{
		_package.Dispose();
	}

	public override string GetRangeFormula(string worksheetName, int row, int column)
	{
		SetCurrentWorksheet(worksheetName);
		return _currentWorksheet.GetFormula(row, column);
	}

	public override object GetRangeValue(string worksheetName, int row, int column)
	{
		SetCurrentWorksheet(worksheetName);
		return _currentWorksheet.GetValue(row, column);
	}

	public override string GetFormat(object value, string format)
	{
		ExcelStyles styles = _package.Workbook.Styles;
		ExcelNumberFormatXml.ExcelFormatTranslator excelFormatTranslator = null;
		foreach (ExcelNumberFormatXml numberFormat in styles.NumberFormats)
		{
			if (numberFormat.Format == format)
			{
				excelFormatTranslator = numberFormat.FormatTranslator;
				break;
			}
		}
		if (excelFormatTranslator == null)
		{
			excelFormatTranslator = new ExcelNumberFormatXml.ExcelFormatTranslator(format, -1);
		}
		return ExcelRangeBase.FormatValue(value, excelFormatTranslator, format, excelFormatTranslator.NetFormat);
	}

	public override List<Token> GetRangeFormulaTokens(string worksheetName, int row, int column)
	{
		return _package.Workbook.Worksheets[worksheetName]._formulaTokens.GetValue(row, column);
	}

	public override bool IsRowHidden(string worksheetName, int row)
	{
		if (_package.Workbook.Worksheets[worksheetName].Row(row).Height != 0.0)
		{
			return _package.Workbook.Worksheets[worksheetName].Row(row).Hidden;
		}
		return true;
	}

	public override void Reset()
	{
		_names = new Dictionary<ulong, INameInfo>();
	}
}
