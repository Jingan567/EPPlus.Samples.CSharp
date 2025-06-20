using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml;
using OfficeOpenXml.Compatibility;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml;

public class ExcelRangeBase : ExcelAddress, IExcelCell, IDisposable, IEnumerable<ExcelRangeBase>, IEnumerable, IEnumerator<ExcelRangeBase>, IEnumerator
{
	private delegate void _changeProp(ExcelRangeBase range, _setValue method, object value);

	private delegate void _setValue(ExcelRangeBase range, object value, int row, int col);

	private class CopiedCell
	{
		internal int Row { get; set; }

		internal int Column { get; set; }

		internal object Value { get; set; }

		internal string Type { get; set; }

		internal object Formula { get; set; }

		internal int? StyleID { get; set; }

		internal Uri HyperLink { get; set; }

		internal ExcelComment Comment { get; set; }

		internal byte Flag { get; set; }
	}

	private struct SortItem<T>
	{
		internal int Row { get; set; }

		internal T[] Items { get; set; }
	}

	private class Comp : IComparer<SortItem<ExcelCoreValue>>
	{
		public int[] columns;

		public bool[] descending;

		public CultureInfo cultureInfo = CultureInfo.CurrentCulture;

		public CompareOptions compareOptions;

		public int Compare(SortItem<ExcelCoreValue> x, SortItem<ExcelCoreValue> y)
		{
			int num = 0;
			for (int i = 0; i < columns.Length; i++)
			{
				object value = x.Items[columns[i]]._value;
				object value2 = y.Items[columns[i]]._value;
				bool flag = ConvertUtil.IsNumeric(value);
				bool flag2 = ConvertUtil.IsNumeric(value2);
				if (flag && flag2)
				{
					double num2 = ConvertUtil.GetValueDouble(value);
					double num3 = ConvertUtil.GetValueDouble(value2);
					if (double.IsNaN(num2))
					{
						num2 = double.MaxValue;
					}
					if (double.IsNaN(num3))
					{
						num3 = double.MaxValue;
					}
					num = ((num2 < num3) ? (-1) : ((num2 > num3) ? 1 : 0));
				}
				else if (!flag && !flag2)
				{
					string strA = ((value == null) ? "" : value.ToString());
					string strB = ((value2 == null) ? "" : value2.ToString());
					num = string.Compare(strA, strB, StringComparison.CurrentCulture);
				}
				else
				{
					num = ((!flag) ? 1 : (-1));
				}
				if (num != 0)
				{
					return num * ((!descending[i]) ? 1 : (-1));
				}
			}
			return 0;
		}
	}

	protected ExcelWorksheet _worksheet;

	internal ExcelWorkbook _workbook;

	private _changeProp _changePropMethod;

	private int _styleID;

	private static _changeProp _setUnknownProp = SetUnknown;

	private static _changeProp _setSingleProp = SetSingle;

	private static _changeProp _setRangeProp = SetRange;

	private static _changeProp _setMultiProp = SetMultiRange;

	private static _setValue _setStyleIdDelegate = Set_StyleID;

	private static _setValue _setValueDelegate = Set_Value;

	private static _setValue _setHyperLinkDelegate = Set_HyperLink;

	private static _setValue _setIsRichTextDelegate = Set_IsRichText;

	private static _setValue _setExistsCommentDelegate = Exists_Comment;

	private static _setValue _setCommentDelegate = Set_Comment;

	protected ExcelRichTextCollection _rtc;

	private CellsStoreEnumerator<ExcelCoreValue> cellEnum;

	private int _enumAddressIx = -1;

	public ExcelStyle Style
	{
		get
		{
			IsRangeValid("styling");
			int styleId = 0;
			if (!_worksheet.ExistsStyleInner(_fromRow, _fromCol, ref styleId) && !_worksheet.ExistsStyleInner(_fromRow, 0, ref styleId))
			{
				styleId = Worksheet.GetColumn(_fromCol)?.StyleID ?? 0;
			}
			return _worksheet.Workbook.Styles.GetStyleObject(styleId, _worksheet.PositionID, base.Address);
		}
	}

	public string StyleName
	{
		get
		{
			IsRangeValid("styling");
			int styleId;
			if (_fromRow == 1 && _toRow == 1048576)
			{
				styleId = GetColumnStyle(_fromCol);
			}
			else if (_fromCol == 1 && _toCol == 16384)
			{
				styleId = 0;
				if (!_worksheet.ExistsStyleInner(_fromRow, 0, ref styleId))
				{
					styleId = GetColumnStyle(_fromCol);
				}
			}
			else
			{
				styleId = 0;
				if (!_worksheet.ExistsStyleInner(_fromRow, _fromCol, ref styleId) && !_worksheet.ExistsStyleInner(_fromRow, 0, ref styleId))
				{
					styleId = GetColumnStyle(_fromCol);
				}
			}
			int num = ((styleId > 0) ? Style.Styles.CellXfs[styleId].XfId : Style.Styles.CellXfs[0].XfId);
			foreach (ExcelNamedStyleXml namedStyle in Style.Styles.NamedStyles)
			{
				if (namedStyle.StyleXfId == num)
				{
					return namedStyle.Name;
				}
			}
			return "";
		}
		set
		{
			_styleID = _worksheet.Workbook.Styles.GetStyleIdFromName(value);
			_ = _fromCol;
			if (_fromRow == 1 && _toRow == 1048576)
			{
				object value2 = _worksheet.GetValue(0, _fromCol);
				ExcelColumn excelColumn = ((value2 != null) ? ((ExcelColumn)value2) : _worksheet.Column(_fromCol));
				excelColumn.StyleName = value;
				excelColumn.StyleID = _styleID;
				CellsStoreEnumerator<ExcelCoreValue> cellsStoreEnumerator = new CellsStoreEnumerator<ExcelCoreValue>(_worksheet._values, 0, _fromCol + 1, 0, _toCol);
				if (cellsStoreEnumerator.Next())
				{
					_ = _fromCol;
					while (excelColumn.ColumnMin <= _toCol)
					{
						if (excelColumn.ColumnMax > _toCol)
						{
							_worksheet.CopyColumn(excelColumn, _toCol + 1, excelColumn.ColumnMax);
							excelColumn.ColumnMax = _toCol;
						}
						excelColumn._styleName = value;
						excelColumn.StyleID = _styleID;
						if (cellsStoreEnumerator.Value._value == null)
						{
							break;
						}
						ExcelColumn excelColumn2 = (ExcelColumn)cellsStoreEnumerator.Value._value;
						if (excelColumn.ColumnMax < excelColumn2.ColumnMax - 1)
						{
							excelColumn.ColumnMax = excelColumn2.ColumnMax - 1;
						}
						excelColumn = excelColumn2;
						cellsStoreEnumerator.Next();
					}
				}
				if (excelColumn.ColumnMax < _toCol)
				{
					excelColumn.ColumnMax = _toCol;
				}
				if (_fromCol == 1 && _toCol == 16384)
				{
					CellsStoreEnumerator<ExcelCoreValue> cellsStoreEnumerator2 = new CellsStoreEnumerator<ExcelCoreValue>(_worksheet._values, 1, 0, 1048576, 0);
					cellsStoreEnumerator2.Next();
					while (cellsStoreEnumerator2.Value._value != null)
					{
						_worksheet.SetStyleInner(cellsStoreEnumerator2.Row, 0, _styleID);
						if (!cellsStoreEnumerator2.Next())
						{
							break;
						}
					}
				}
			}
			else if (_fromCol == 1 && _toCol == 16384)
			{
				for (int i = _fromRow; i <= _toRow; i++)
				{
					_worksheet.Row(i)._styleName = value;
					_worksheet.Row(i).StyleID = _styleID;
				}
			}
			if ((_fromRow != 1 || _toRow != 1048576) && (_fromCol != 1 || _toCol != 16384))
			{
				for (int j = _fromCol; j <= _toCol; j++)
				{
					for (int k = _fromRow; k <= _toRow; k++)
					{
						_worksheet.SetStyleInner(k, j, _styleID);
					}
				}
			}
			else
			{
				CellsStoreEnumerator<ExcelCoreValue> cellsStoreEnumerator3 = new CellsStoreEnumerator<ExcelCoreValue>(_worksheet._values, _fromRow, _fromCol, _toRow, _toCol);
				while (cellsStoreEnumerator3.Next())
				{
					_worksheet.SetStyleInner(cellsStoreEnumerator3.Row, cellsStoreEnumerator3.Column, _styleID);
				}
			}
		}
	}

	public int StyleID
	{
		get
		{
			int styleId = 0;
			if (!_worksheet.ExistsStyleInner(_fromRow, _fromCol, ref styleId) && !_worksheet.ExistsStyleInner(_fromRow, 0, ref styleId))
			{
				styleId = _worksheet.GetStyleInner(0, _fromCol);
			}
			return styleId;
		}
		set
		{
			_changePropMethod(this, _setStyleIdDelegate, value);
		}
	}

	public object Value
	{
		get
		{
			if (base.IsName)
			{
				if (_worksheet == null)
				{
					return _workbook._names[_address].NameValue;
				}
				return _worksheet.Names[_address].NameValue;
			}
			if (_fromRow == _toRow && _fromCol == _toCol)
			{
				return _worksheet.GetValue(_fromRow, _fromCol);
			}
			return GetValueArray();
		}
		set
		{
			if (base.IsName)
			{
				if (_worksheet == null)
				{
					_workbook._names[_address].NameValue = value;
				}
				else
				{
					_worksheet.Names[_address].NameValue = value;
				}
			}
			else
			{
				_changePropMethod(this, _setValueDelegate, value);
			}
		}
	}

	public string Text => GetFormattedText(forWidthCalc: false);

	internal string TextForWidth => GetFormattedText(forWidthCalc: true);

	public string Formula
	{
		get
		{
			if (base.IsName)
			{
				if (_worksheet == null)
				{
					return _workbook._names[_address].NameFormula;
				}
				return _worksheet.Names[_address].NameFormula;
			}
			return _worksheet.GetFormula(_fromRow, _fromCol);
		}
		set
		{
			if (base.IsName)
			{
				if (_worksheet == null)
				{
					_workbook._names[_address].NameFormula = value;
				}
				else
				{
					_worksheet.Names[_address].NameFormula = value;
				}
				return;
			}
			if (value == null || value.Trim() == "")
			{
				Value = null;
				return;
			}
			if (_fromRow == _toRow && _fromCol == _toCol)
			{
				Set_Formula(this, value, _fromRow, _fromCol);
				return;
			}
			Set_SharedFormula(this, value, this, IsArray: false);
			if (Addresses == null)
			{
				return;
			}
			foreach (ExcelAddress address in Addresses)
			{
				Set_SharedFormula(this, value, address, IsArray: false);
			}
		}
	}

	public string FormulaR1C1
	{
		get
		{
			IsRangeValid("FormulaR1C1");
			return _worksheet.GetFormulaR1C1(_fromRow, _fromCol);
		}
		set
		{
			IsRangeValid("FormulaR1C1");
			if (value.Length > 0 && value[0] == '=')
			{
				value = value.Substring(1, value.Length - 1);
			}
			if (value == null || value.Trim() == "")
			{
				_worksheet.Cells[ExcelCellBase.TranslateFromR1C1(value, _fromRow, _fromCol)].Value = null;
				return;
			}
			if (Addresses == null)
			{
				Set_SharedFormula(this, ExcelCellBase.TranslateFromR1C1(value, _fromRow, _fromCol), this, IsArray: false);
				return;
			}
			Set_SharedFormula(this, ExcelCellBase.TranslateFromR1C1(value, _fromRow, _fromCol), new ExcelAddress(base.WorkSheet, base.FirstAddress), IsArray: false);
			foreach (ExcelAddress address in Addresses)
			{
				Set_SharedFormula(this, ExcelCellBase.TranslateFromR1C1(value, address.Start.Row, address.Start.Column), address, IsArray: false);
			}
		}
	}

	public Uri Hyperlink
	{
		get
		{
			IsRangeValid("formulaR1C1");
			return _worksheet._hyperLinks.GetValue(_fromRow, _fromCol);
		}
		set
		{
			_changePropMethod(this, _setHyperLinkDelegate, value);
		}
	}

	public bool Merge
	{
		get
		{
			IsRangeValid("merging");
			for (int i = _fromCol; i <= _toCol; i++)
			{
				for (int j = _fromRow; j <= _toRow; j++)
				{
					if (_worksheet.MergedCells[j, i] == null)
					{
						return false;
					}
				}
			}
			return true;
		}
		set
		{
			IsRangeValid("merging");
			_worksheet.MergedCells.Clear(this);
			if (value)
			{
				_worksheet.MergedCells.Add(new ExcelAddressBase(base.FirstAddress), doValidate: true);
				if (Addresses == null)
				{
					return;
				}
				{
					foreach (ExcelAddress address in Addresses)
					{
						_worksheet.MergedCells.Clear(address);
						_worksheet.MergedCells.Add(address, doValidate: true);
					}
					return;
				}
			}
			if (Addresses == null)
			{
				return;
			}
			foreach (ExcelAddress address2 in Addresses)
			{
				_worksheet.MergedCells.Clear(address2);
			}
		}
	}

	public bool AutoFilter
	{
		get
		{
			IsRangeValid("autofilter");
			ExcelAddressBase autoFilterAddress = _worksheet.AutoFilterAddress;
			if (autoFilterAddress == null)
			{
				return false;
			}
			if (_fromRow >= autoFilterAddress.Start.Row && _toRow <= autoFilterAddress.End.Row && _fromCol >= autoFilterAddress.Start.Column && _toCol <= autoFilterAddress.End.Column)
			{
				return true;
			}
			return false;
		}
		set
		{
			IsRangeValid("autofilter");
			if (_worksheet.AutoFilterAddress != null)
			{
				eAddressCollition eAddressCollition = Collide(_worksheet.AutoFilterAddress);
				if (!value && (eAddressCollition == eAddressCollition.Partly || eAddressCollition == eAddressCollition.No))
				{
					throw new InvalidOperationException("Can't remote Autofilter. Current autofilter does not match selected range.");
				}
			}
			if (_worksheet.Names.ContainsKey("_xlnm._FilterDatabase"))
			{
				_worksheet.Names.Remove("_xlnm._FilterDatabase");
			}
			if (value)
			{
				_worksheet.AutoFilterAddress = this;
				_worksheet.Names.Add("_xlnm._FilterDatabase", this).IsNameHidden = true;
			}
			else
			{
				_worksheet.AutoFilterAddress = null;
			}
		}
	}

	public bool IsRichText
	{
		get
		{
			IsRangeValid("richtext");
			return _worksheet._flags.GetFlagValue(_fromRow, _fromCol, CellFlags.RichText);
		}
		set
		{
			_changePropMethod(this, _setIsRichTextDelegate, value);
		}
	}

	public bool IsArrayFormula
	{
		get
		{
			IsRangeValid("arrayformulas");
			return _worksheet._flags.GetFlagValue(_fromRow, _fromCol, CellFlags.ArrayFormula);
		}
	}

	public ExcelRichTextCollection RichText
	{
		get
		{
			IsRangeValid("richtext");
			if (_rtc == null)
			{
				_rtc = GetRichText(_fromRow, _fromCol);
			}
			return _rtc;
		}
	}

	public ExcelComment Comment
	{
		get
		{
			IsRangeValid("comments");
			int value = -1;
			if (_worksheet.Comments.Count > 0 && _worksheet._commentsStore.Exists(_fromRow, _fromCol, ref value))
			{
				return _worksheet._comments[value];
			}
			return null;
		}
	}

	public ExcelWorksheet Worksheet => _worksheet;

	public new string FullAddress
	{
		get
		{
			string text = ExcelCellBase.GetFullAddress(_worksheet.Name, _address);
			if (Addresses != null)
			{
				foreach (ExcelAddress address in Addresses)
				{
					text = text + "," + ExcelCellBase.GetFullAddress(_worksheet.Name, address.Address);
				}
			}
			return text;
		}
	}

	public string FullAddressAbsolute
	{
		get
		{
			string worksheetName = (string.IsNullOrEmpty(_wb) ? _ws : ("[" + _wb.Replace("'", "''") + "]" + _ws));
			string text;
			if (Addresses == null)
			{
				text = ExcelCellBase.GetFullAddress(worksheetName, ExcelCellBase.GetAddress(_fromRow, _fromCol, _toRow, _toCol, Absolute: true));
			}
			else
			{
				text = "";
				foreach (ExcelAddress address in Addresses)
				{
					if (text != "")
					{
						text += ",";
					}
					text = ((!(address.Address == "#REF!")) ? (text + ExcelCellBase.GetFullAddress(worksheetName, ExcelCellBase.GetAddress(address.Start.Row, address.Start.Column, address.End.Row, address.End.Column, Absolute: true))) : (text + ExcelCellBase.GetFullAddress(worksheetName, "#REF!")));
				}
			}
			return text;
		}
	}

	internal string FullAddressAbsoluteNoFullRowCol
	{
		get
		{
			string worksheetName = (string.IsNullOrEmpty(_wb) ? _ws : ("[" + _wb.Replace("'", "''") + "]" + _ws));
			string text;
			if (Addresses == null)
			{
				text = ExcelCellBase.GetFullAddress(worksheetName, ExcelCellBase.GetAddress(_fromRow, _fromCol, _toRow, _toCol, Absolute: true), fullRowCol: false);
			}
			else
			{
				text = "";
				foreach (ExcelAddress address in Addresses)
				{
					if (text != "")
					{
						text += ",";
					}
					text += ExcelCellBase.GetFullAddress(worksheetName, ExcelCellBase.GetAddress(address.Start.Row, address.Start.Column, address.End.Row, address.End.Column, Absolute: true), fullRowCol: false);
				}
			}
			return text;
		}
	}

	public IRangeConditionalFormatting ConditionalFormatting => new RangeConditionalFormatting(_worksheet, new ExcelAddress(base.Address));

	public IRangeDataValidation DataValidation => new RangeDataValidation(_worksheet, base.Address);

	public ExcelRangeBase Current => new ExcelRangeBase(_worksheet, ExcelCellBase.GetAddress(cellEnum.Row, cellEnum.Column));

	object IEnumerator.Current => new ExcelRangeBase(_worksheet, ExcelCellBase.GetAddress(cellEnum.Row, cellEnum.Column));

	internal ExcelRangeBase(ExcelWorksheet xlWorksheet)
	{
		_worksheet = xlWorksheet;
		_ws = _worksheet.Name;
		_workbook = _worksheet.Workbook;
		SetDelegate();
	}

	protected internal override void ChangeAddress()
	{
		if (base.Table != null)
		{
			SetRCFromTable(_workbook._package, null);
		}
		SetDelegate();
	}

	internal ExcelRangeBase(ExcelWorksheet xlWorksheet, string address)
		: base((xlWorksheet == null) ? "" : xlWorksheet.Name, address)
	{
		_worksheet = xlWorksheet;
		_workbook = _worksheet.Workbook;
		SetRCFromTable(_worksheet._package, null);
		if (string.IsNullOrEmpty(_ws))
		{
			_ws = ((_worksheet == null) ? "" : _worksheet.Name);
		}
		SetDelegate();
	}

	internal ExcelRangeBase(ExcelWorkbook wb, ExcelWorksheet xlWorksheet, string address, bool isName)
		: base((xlWorksheet == null) ? "" : xlWorksheet.Name, address, isName)
	{
		SetRCFromTable(wb._package, null);
		_worksheet = xlWorksheet;
		_workbook = wb;
		if (string.IsNullOrEmpty(_ws))
		{
			_ws = xlWorksheet?.Name;
		}
		SetDelegate();
	}

	private void SetDelegate()
	{
		if (_fromRow == -1)
		{
			_changePropMethod = SetUnknown;
		}
		else if (_fromRow == _toRow && _fromCol == _toCol && Addresses == null)
		{
			_changePropMethod = SetSingle;
		}
		else if (Addresses == null)
		{
			_changePropMethod = SetRange;
		}
		else
		{
			_changePropMethod = SetMultiRange;
		}
	}

	private static void SetUnknown(ExcelRangeBase range, _setValue valueMethod, object value)
	{
		if (range._fromRow == -1)
		{
			range.SetToSelectedRange();
		}
		range.SetDelegate();
		range._changePropMethod(range, valueMethod, value);
	}

	private static void SetSingle(ExcelRangeBase range, _setValue valueMethod, object value)
	{
		valueMethod(range, value, range._fromRow, range._fromCol);
	}

	private static void SetRange(ExcelRangeBase range, _setValue valueMethod, object value)
	{
		range.SetValueAddress(range, valueMethod, value);
	}

	private static void SetMultiRange(ExcelRangeBase range, _setValue valueMethod, object value)
	{
		range.SetValueAddress(range, valueMethod, value);
		foreach (ExcelAddress address in range.Addresses)
		{
			range.SetValueAddress(address, valueMethod, value);
		}
	}

	private void SetValueAddress(ExcelAddress address, _setValue valueMethod, object value)
	{
		IsRangeValid("");
		if (_fromRow == 1 && _fromCol == 1 && _toRow == 1048576 && _toCol == 16384)
		{
			throw new ArgumentException("Can't reference all cells. Please use the indexer to set the range");
		}
		if (value is object[,] && valueMethod == new _setValue(Set_Value))
		{
			_worksheet.SetRangeValueInner(address.Start.Row, address.Start.Column, address.End.Row, address.End.Column, (object[,])value);
			return;
		}
		for (int i = address.Start.Column; i <= address.End.Column; i++)
		{
			for (int j = address.Start.Row; j <= address.End.Row; j++)
			{
				valueMethod(this, value, j, i);
			}
		}
	}

	private static void Set_StyleID(ExcelRangeBase range, object value, int row, int col)
	{
		range._worksheet.SetStyleInner(row, col, (int)value);
	}

	private static void Set_StyleName(ExcelRangeBase range, object value, int row, int col)
	{
		range._worksheet.SetStyleInner(row, col, range._styleID);
	}

	private static void Set_Value(ExcelRangeBase range, object value, int row, int col)
	{
		object value2 = range._worksheet._formulas.GetValue(row, col);
		if (value2 is int)
		{
			range.SplitFormulas(range._worksheet.Cells[row, col]);
		}
		if (value2 != null)
		{
			range._worksheet._formulas.SetValue(row, col, string.Empty);
		}
		range._worksheet.SetValueInner(row, col, value);
	}

	private static void Set_Formula(ExcelRangeBase range, object value, int row, int col)
	{
		object value2 = range._worksheet._formulas.GetValue(row, col);
		if (value2 is int && (int)value2 >= 0)
		{
			range.SplitFormulas(range._worksheet.Cells[row, col]);
		}
		string text = ((value == null) ? string.Empty : value.ToString());
		if (text == string.Empty)
		{
			range._worksheet._formulas.SetValue(row, col, string.Empty);
			return;
		}
		if (text[0] == '=')
		{
			value = text.Substring(1, text.Length - 1);
		}
		range._worksheet._formulas.SetValue(row, col, text);
		range._worksheet.SetValueInner(row, col, null);
	}

	private static void Set_SharedFormula(ExcelRangeBase range, string value, ExcelAddress address, bool IsArray)
	{
		if (range._fromRow == 1 && range._fromCol == 1 && range._toRow == 1048576 && range._toCol == 16384)
		{
			throw new InvalidOperationException("Can't set a formula for the entire worksheet");
		}
		if (address.Start.Row == address.End.Row && address.Start.Column == address.End.Column && !IsArray)
		{
			Set_Formula(range, value, address.Start.Row, address.Start.Column);
			return;
		}
		range.CheckAndSplitSharedFormula(address);
		ExcelWorksheet.Formulas formulas = new ExcelWorksheet.Formulas(SourceCodeTokenizer.Default);
		formulas.Formula = value;
		formulas.Index = range._worksheet.GetMaxShareFunctionIndex(IsArray);
		formulas.Address = address.FirstAddress;
		formulas.StartCol = address.Start.Column;
		formulas.StartRow = address.Start.Row;
		formulas.IsArray = IsArray;
		range._worksheet._sharedFormulas.Add(formulas.Index, formulas);
		for (int i = address.Start.Column; i <= address.End.Column; i++)
		{
			for (int j = address.Start.Row; j <= address.End.Row; j++)
			{
				range._worksheet._formulas.SetValue(j, i, formulas.Index);
				range._worksheet._flags.SetFlagValue(j, i, value: true, CellFlags.ArrayFormula);
				range._worksheet.SetValueInner(j, i, null);
			}
		}
	}

	private static void Set_HyperLink(ExcelRangeBase range, object value, int row, int col)
	{
		if (value is Uri)
		{
			range._worksheet._hyperLinks.SetValue(row, col, (Uri)value);
			if (value is ExcelHyperLink)
			{
				range._worksheet.SetValueInner(row, col, ((ExcelHyperLink)value).Display);
				return;
			}
			object valueInner = range._worksheet.GetValueInner(row, col);
			if (valueInner == null || valueInner.ToString() == "")
			{
				range._worksheet.SetValueInner(row, col, ((Uri)value).OriginalString);
			}
		}
		else
		{
			range._worksheet._hyperLinks.SetValue(row, col, null);
			range._worksheet.SetValueInner(row, col, null);
		}
	}

	private static void Set_IsRichText(ExcelRangeBase range, object value, int row, int col)
	{
		range._worksheet._flags.SetFlagValue(row, col, (bool)value, CellFlags.RichText);
	}

	private static void Exists_Comment(ExcelRangeBase range, object value, int row, int col)
	{
		if (range._worksheet._commentsStore.Exists(row, col))
		{
			throw new InvalidOperationException($"Cell {new ExcelCellAddress(row, col).Address} already contain a comment.");
		}
	}

	private static void Set_Comment(ExcelRangeBase range, object value, int row, int col)
	{
		string[] array = (string[])value;
		range._worksheet.Comments.Add(new ExcelRangeBase(range._worksheet, ExcelCellBase.GetAddress(range._fromRow, range._fromCol)), array[0], array[1]);
	}

	private void SetToSelectedRange()
	{
		if (_worksheet.View.SelectedRange == "")
		{
			base.Address = "A1";
		}
		else
		{
			base.Address = _worksheet.View.SelectedRange;
		}
	}

	private void IsRangeValid(string type)
	{
		if (_fromRow > 0)
		{
			return;
		}
		if (_address == "")
		{
			SetToSelectedRange();
			return;
		}
		if (type == "")
		{
			throw new InvalidOperationException($"Range is not valid for this operation: {_address}");
		}
		throw new InvalidOperationException($"Range is not valid for {type} : {_address}");
	}

	internal void UpdateAddress(string address)
	{
		throw new NotImplementedException();
	}

	private int GetColumnStyle(int col)
	{
		object value = null;
		if (_worksheet.ExistsValueInner(0, col, ref value))
		{
			return (value as ExcelColumn).StyleID;
		}
		int row = 0;
		if (_worksheet._values.PrevCell(ref row, ref col) && (_worksheet.GetValueInner(row, col) as ExcelColumn).ColumnMax >= col)
		{
			return _worksheet.GetStyleInner(row, col);
		}
		return 0;
	}

	private bool IsInfinityValue(object value)
	{
		double? num = value as double?;
		if (num.HasValue && (double.IsNegativeInfinity(num.Value) || double.IsPositiveInfinity(num.Value)))
		{
			return true;
		}
		return false;
	}

	private object GetValueArray()
	{
		ExcelAddressBase excelAddressBase;
		if (_fromRow == 1 && _fromCol == 1 && _toRow == 1048576 && _toCol == 16384)
		{
			excelAddressBase = _worksheet.Dimension;
			if (excelAddressBase == null)
			{
				return null;
			}
		}
		else
		{
			excelAddressBase = this;
		}
		object[,] array = new object[excelAddressBase._toRow - excelAddressBase._fromRow + 1, excelAddressBase._toCol - excelAddressBase._fromCol + 1];
		for (int i = excelAddressBase._fromCol; i <= excelAddressBase._toCol; i++)
		{
			for (int j = excelAddressBase._fromRow; j <= excelAddressBase._toRow; j++)
			{
				object value = null;
				if (_worksheet.ExistsValueInner(j, i, ref value))
				{
					if (_worksheet._flags.GetFlagValue(j, i, CellFlags.RichText))
					{
						array[j - excelAddressBase._fromRow, i - excelAddressBase._fromCol] = GetRichText(j, i).Text;
					}
					else
					{
						array[j - excelAddressBase._fromRow, i - excelAddressBase._fromCol] = value;
					}
				}
			}
		}
		return array;
	}

	private ExcelAddressBase GetAddressDim(ExcelRangeBase addr)
	{
		ExcelAddressBase dimension = _worksheet.Dimension;
		int num = ((addr._fromRow < dimension._fromRow) ? dimension._fromRow : addr._fromRow);
		int num2 = ((addr._fromCol < dimension._fromCol) ? dimension._fromCol : addr._fromCol);
		int num3 = ((addr._toRow > dimension._toRow) ? dimension._toRow : addr._toRow);
		int toColumn = ((addr._toCol > dimension._toCol) ? dimension._toCol : addr._toCol);
		if (addr._fromRow == num && addr._fromCol == num2 && addr._toRow == num3 && addr._toCol == _toCol)
		{
			return addr;
		}
		if (_fromRow > _toRow || _fromCol > _toCol)
		{
			return null;
		}
		return new ExcelAddressBase(num, num2, num3, toColumn);
	}

	private object GetSingleValue()
	{
		if (IsRichText)
		{
			return RichText.Text;
		}
		return _worksheet.GetValueInner(_fromRow, _fromCol);
	}

	public void AutoFitColumns()
	{
		AutoFitColumns(_worksheet.DefaultColWidth);
	}

	public void AutoFitColumns(double MinimumWidth)
	{
		AutoFitColumns(MinimumWidth, double.MaxValue);
	}

	public void AutoFitColumns(double MinimumWidth, double MaximumWidth)
	{
		if (_worksheet.Dimension == null)
		{
			return;
		}
		if (_fromCol < 1 || _fromRow < 1)
		{
			SetToSelectedRange();
		}
		Dictionary<int, Font> dictionary = new Dictionary<int, Font>();
		bool doAdjustDrawings = _worksheet._package.DoAdjustDrawings;
		_worksheet._package.DoAdjustDrawings = false;
		int[,] drawingWidths = _worksheet.Drawings.GetDrawingWidths();
		int num = ((_fromCol > _worksheet.Dimension._fromCol) ? _fromCol : _worksheet.Dimension._fromCol);
		int num2 = ((_toCol < _worksheet.Dimension._toCol) ? _toCol : _worksheet.Dimension._toCol);
		if (num > num2)
		{
			return;
		}
		if (Addresses == null)
		{
			SetMinWidth(MinimumWidth, num, num2);
		}
		else
		{
			foreach (ExcelAddress address in Addresses)
			{
				num = ((address._fromCol > _worksheet.Dimension._fromCol) ? address._fromCol : _worksheet.Dimension._fromCol);
				num2 = ((address._toCol < _worksheet.Dimension._toCol) ? address._toCol : _worksheet.Dimension._toCol);
				SetMinWidth(MinimumWidth, num, num2);
			}
		}
		List<ExcelAddressBase> list = new List<ExcelAddressBase>();
		if (_worksheet.AutoFilterAddress != null)
		{
			list.Add(new ExcelAddressBase(_worksheet.AutoFilterAddress._fromRow, _worksheet.AutoFilterAddress._fromCol, _worksheet.AutoFilterAddress._fromRow, _worksheet.AutoFilterAddress._toCol));
			list[list.Count - 1]._ws = base.WorkSheet;
		}
		foreach (ExcelTable table in _worksheet.Tables)
		{
			if (table.AutoFilterAddress != null)
			{
				list.Add(new ExcelAddressBase(table.AutoFilterAddress._fromRow, table.AutoFilterAddress._fromCol, table.AutoFilterAddress._fromRow, table.AutoFilterAddress._toCol));
				list[list.Count - 1]._ws = base.WorkSheet;
			}
		}
		ExcelStyles styles = _worksheet.Workbook.Styles;
		ExcelFontXml excelFontXml = styles.Fonts[styles.CellXfs[0].FontId];
		FontStyle fontStyle = FontStyle.Regular;
		if (excelFontXml.Bold)
		{
			fontStyle |= FontStyle.Bold;
		}
		if (excelFontXml.UnderLine)
		{
			fontStyle |= FontStyle.Underline;
		}
		if (excelFontXml.Italic)
		{
			fontStyle |= FontStyle.Italic;
		}
		if (excelFontXml.Strike)
		{
			fontStyle |= FontStyle.Strikeout;
		}
		new Font(excelFontXml.Name, excelFontXml.Size, fontStyle);
		float num3 = Convert.ToSingle(ExcelWorkbook.GetWidthPixels(excelFontXml.Name, excelFontXml.Size));
		Graphics graphics = null;
		try
		{
			graphics = Graphics.FromImage(new Bitmap(1, 1));
			graphics.PageUnit = GraphicsUnit.Pixel;
		}
		catch
		{
			return;
		}
		using (IEnumerator<ExcelRangeBase> enumerator3 = GetEnumerator())
		{
			while (enumerator3.MoveNext())
			{
				ExcelRangeBase current3 = enumerator3.Current;
				if (_worksheet.Column(current3.Start.Column).Hidden || current3.Merge || current3.Style.WrapText)
				{
					continue;
				}
				int fontId = styles.CellXfs[current3.StyleID].FontId;
				Font font;
				if (dictionary.ContainsKey(fontId))
				{
					font = dictionary[fontId];
				}
				else
				{
					ExcelFontXml excelFontXml2 = styles.Fonts[fontId];
					fontStyle = FontStyle.Regular;
					if (excelFontXml2.Bold)
					{
						fontStyle |= FontStyle.Bold;
					}
					if (excelFontXml2.UnderLine)
					{
						fontStyle |= FontStyle.Underline;
					}
					if (excelFontXml2.Italic)
					{
						fontStyle |= FontStyle.Italic;
					}
					if (excelFontXml2.Strike)
					{
						fontStyle |= FontStyle.Strikeout;
					}
					font = new Font(excelFontXml2.Name, excelFontXml2.Size, fontStyle);
					dictionary.Add(fontId, font);
				}
				int indent = styles.CellXfs[current3.StyleID].Indent;
				string textForWidth = current3.TextForWidth;
				string text = textForWidth + ((indent > 0 && !string.IsNullOrEmpty(textForWidth)) ? new string('_', indent) : "");
				SizeF sizeF = graphics.MeasureString(text, font, 10000, StringFormat.GenericDefault);
				double num4 = styles.CellXfs[current3.StyleID].TextRotation;
				double num5;
				if (num4 <= 0.0)
				{
					num5 = (sizeF.Width + 5f) / num3;
				}
				else
				{
					num4 = ((num4 <= 90.0) ? num4 : (num4 - 90.0));
					num5 = ((double)(sizeF.Width - sizeF.Height) * Math.Abs(Math.Cos(Math.PI * num4 / 180.0)) + (double)sizeF.Height + 5.0) / (double)num3;
				}
				foreach (ExcelAddressBase item in list)
				{
					if (item.Collide(current3) != 0)
					{
						num5 += 2.25;
						break;
					}
				}
				if (num5 > _worksheet.Column(current3._fromCol).Width)
				{
					_worksheet.Column(current3._fromCol).Width = ((num5 > MaximumWidth) ? MaximumWidth : num5);
				}
			}
		}
		_worksheet.Drawings.AdjustWidth(drawingWidths);
		_worksheet._package.DoAdjustDrawings = doAdjustDrawings;
	}

	private void SetMinWidth(double minimumWidth, int fromCol, int toCol)
	{
		CellsStoreEnumerator<ExcelCoreValue> cellsStoreEnumerator = new CellsStoreEnumerator<ExcelCoreValue>(_worksheet._values, 0, fromCol, 0, toCol);
		int num = fromCol;
		foreach (ExcelCoreValue item in cellsStoreEnumerator)
		{
			ExcelColumn excelColumn = (ExcelColumn)item._value;
			excelColumn.Width = minimumWidth;
			if (_worksheet.DefaultColWidth > minimumWidth && excelColumn.ColumnMin > num)
			{
				ExcelColumn excelColumn2 = _worksheet.Column(num);
				excelColumn2.ColumnMax = excelColumn.ColumnMin - 1;
				excelColumn2.Width = minimumWidth;
			}
			num = excelColumn.ColumnMax + 1;
		}
		if (_worksheet.DefaultColWidth > minimumWidth && num < toCol)
		{
			ExcelColumn excelColumn3 = _worksheet.Column(num);
			excelColumn3.ColumnMax = toCol;
			excelColumn3.Width = minimumWidth;
		}
	}

	private string GetFormattedText(bool forWidthCalc)
	{
		object value = Value;
		if (value == null)
		{
			return "";
		}
		ExcelStyles styles = Worksheet.Workbook.Styles;
		int numberFormatId = styles.CellXfs[StyleID].NumberFormatId;
		ExcelNumberFormatXml.ExcelFormatTranslator excelFormatTranslator = null;
		for (int i = 0; i < styles.NumberFormats.Count; i++)
		{
			if (numberFormatId == styles.NumberFormats[i].NumFmtId)
			{
				excelFormatTranslator = styles.NumberFormats[i].FormatTranslator;
				break;
			}
		}
		if (excelFormatTranslator == null)
		{
			excelFormatTranslator = styles.NumberFormats[0].FormatTranslator;
		}
		string format;
		string textFormat;
		if (forWidthCalc)
		{
			format = excelFormatTranslator.NetFormatForWidth;
			textFormat = excelFormatTranslator.NetTextFormatForWidth;
		}
		else
		{
			format = excelFormatTranslator.NetFormat;
			textFormat = excelFormatTranslator.NetTextFormat;
		}
		return FormatValue(value, excelFormatTranslator, format, textFormat);
	}

	internal static string FormatValue(object v, ExcelNumberFormatXml.ExcelFormatTranslator nf, string format, string textFormat)
	{
		if (v is decimal || TypeCompat.IsPrimitive(v))
		{
			double d;
			try
			{
				d = Convert.ToDouble(v);
			}
			catch
			{
				return "";
			}
			if (nf.DataType == ExcelNumberFormatXml.eFormatType.Number)
			{
				if (string.IsNullOrEmpty(nf.FractionFormat))
				{
					return d.ToString(format, nf.Culture);
				}
				return nf.FormatFraction(d);
			}
			if (nf.DataType == ExcelNumberFormatXml.eFormatType.DateTime)
			{
				return GetDateText(DateTime.FromOADate(d), format, nf.Culture);
			}
			return v.ToString();
		}
		if (v is DateTime)
		{
			if (nf.DataType == ExcelNumberFormatXml.eFormatType.DateTime)
			{
				return GetDateText((DateTime)v, format, nf.Culture);
			}
			double d2 = ((DateTime)v).ToOADate();
			if (string.IsNullOrEmpty(nf.FractionFormat))
			{
				return d2.ToString(format, nf.Culture);
			}
			return nf.FormatFraction(d2);
		}
		if (v is TimeSpan)
		{
			if (nf.DataType == ExcelNumberFormatXml.eFormatType.DateTime)
			{
				return GetDateText(new DateTime(((TimeSpan)v).Ticks), format, nf.Culture);
			}
			double d3 = new DateTime(0L).Add((TimeSpan)v).ToOADate();
			if (string.IsNullOrEmpty(nf.FractionFormat))
			{
				return d3.ToString(format, nf.Culture);
			}
			return nf.FormatFraction(d3);
		}
		if (textFormat == "")
		{
			return v.ToString();
		}
		return string.Format(textFormat, v);
	}

	private static string GetDateText(DateTime d, string format, CultureInfo culture)
	{
		switch (format)
		{
		case "d":
		case "D":
			return d.Day.ToString();
		case "M":
			return d.Month.ToString();
		case "m":
			return d.Minute.ToString();
		default:
			if (format.ToLower() == "y" || format.ToLower() == "yy")
			{
				return d.ToString("yy", culture);
			}
			if (format.ToLower() == "yyy" || format.ToLower() == "yyyy")
			{
				return d.ToString("yyy", culture);
			}
			return d.ToString(format, culture);
		}
	}

	private ExcelRichTextCollection GetRichText(int row, int col)
	{
		XmlDocument xmlDocument = new XmlDocument();
		object valueInner = _worksheet.GetValueInner(row, col);
		bool flagValue = _worksheet._flags.GetFlagValue(row, col, CellFlags.RichText);
		if (valueInner != null)
		{
			if (flagValue)
			{
				XmlHelper.LoadXmlSafe(xmlDocument, "<d:si xmlns:d=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" >" + valueInner.ToString() + "</d:si>", Encoding.UTF8);
			}
			else
			{
				xmlDocument.LoadXml("<d:si xmlns:d=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ><d:r><d:t>" + ConvertUtil.ExcelEscapeString(valueInner.ToString()) + "</d:t></d:r></d:si>");
			}
		}
		else
		{
			xmlDocument.LoadXml("<d:si xmlns:d=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" />");
		}
		return new ExcelRichTextCollection(_worksheet.NameSpaceManager, xmlDocument.SelectSingleNode("d:si", _worksheet.NameSpaceManager), this);
	}

	internal void SetValueRichText(object value)
	{
		if (_fromRow == 1 && _fromCol == 1 && _toRow == 1048576 && _toCol == 16384)
		{
			SetValue(value, 1, 1);
		}
		else
		{
			SetValue(value, _fromRow, _fromCol);
		}
	}

	private void SetValue(object value, int row, int col)
	{
		_worksheet.SetValue(row, col, value);
		_worksheet._formulas.SetValue(row, col, "");
	}

	internal void SetSharedFormulaID(int id)
	{
		for (int i = _fromCol; i <= _toCol; i++)
		{
			for (int j = _fromRow; j <= _toRow; j++)
			{
				_worksheet._formulas.SetValue(j, i, id);
			}
		}
	}

	private void CheckAndSplitSharedFormula(ExcelAddressBase address)
	{
		for (int i = address._fromCol; i <= address._toCol; i++)
		{
			for (int j = address._fromRow; j <= address._toRow; j++)
			{
				object value = _worksheet._formulas.GetValue(j, i);
				if (value is int && (int)value >= 0)
				{
					SplitFormulas(address);
					return;
				}
			}
		}
	}

	private void SplitFormulas(ExcelAddressBase address)
	{
		List<int> list = new List<int>();
		for (int i = address._fromCol; i <= address._toCol; i++)
		{
			for (int j = address._fromRow; j <= address._toRow; j++)
			{
				if (_worksheet._formulas.GetValue(j, i) is int num && num >= 0 && !list.Contains(num))
				{
					if (_worksheet._sharedFormulas[num].IsArray && Collide(_worksheet.Cells[_worksheet._sharedFormulas[num].Address]) == eAddressCollition.Partly)
					{
						throw new InvalidOperationException("Can not overwrite a part of an array-formula");
					}
					list.Add(num);
				}
			}
		}
		foreach (int item in list)
		{
			SplitFormula(address, item);
		}
	}

	private void SplitFormula(ExcelAddressBase address, int ix)
	{
		ExcelWorksheet.Formulas formulas = _worksheet._sharedFormulas[ix];
		ExcelRange excelRange = _worksheet.Cells[formulas.Address];
		eAddressCollition eAddressCollition = address.Collide(excelRange);
		if (eAddressCollition == eAddressCollition.Equal || eAddressCollition == eAddressCollition.Inside)
		{
			_worksheet._sharedFormulas.Remove(ix);
			return;
		}
		eAddressCollition eAddressCollition2 = address.Collide(new ExcelAddressBase(excelRange._fromRow, excelRange._fromCol, excelRange._fromRow, excelRange._fromCol));
		if (eAddressCollition != eAddressCollition.Partly || (eAddressCollition2 != eAddressCollition.Inside && eAddressCollition2 != eAddressCollition.Equal))
		{
			return;
		}
		bool flag = false;
		string formulaR1C = excelRange.FormulaR1C1;
		if (excelRange._fromRow < _fromRow)
		{
			formulas.Address = ExcelCellBase.GetAddress(excelRange._fromRow, excelRange._fromCol, _fromRow - 1, excelRange._toCol);
			flag = true;
		}
		if (excelRange._fromCol < address._fromCol)
		{
			if (flag)
			{
				formulas = new ExcelWorksheet.Formulas(SourceCodeTokenizer.Default);
				formulas.Index = _worksheet.GetMaxShareFunctionIndex(isArray: false);
				formulas.StartCol = excelRange._fromCol;
				formulas.IsArray = false;
				_worksheet._sharedFormulas.Add(formulas.Index, formulas);
			}
			else
			{
				flag = true;
			}
			if (excelRange._fromRow < address._fromRow)
			{
				formulas.StartRow = address._fromRow;
			}
			else
			{
				formulas.StartRow = excelRange._fromRow;
			}
			if (excelRange._toRow < address._toRow)
			{
				formulas.Address = ExcelCellBase.GetAddress(formulas.StartRow, formulas.StartCol, excelRange._toRow, address._fromCol - 1);
			}
			else
			{
				formulas.Address = ExcelCellBase.GetAddress(formulas.StartRow, formulas.StartCol, address._toRow, address._fromCol - 1);
			}
			formulas.Formula = ExcelCellBase.TranslateFromR1C1(formulaR1C, formulas.StartRow, formulas.StartCol);
			_worksheet.Cells[formulas.Address].SetSharedFormulaID(formulas.Index);
		}
		if (excelRange._toCol > address._toCol)
		{
			if (flag)
			{
				formulas = new ExcelWorksheet.Formulas(SourceCodeTokenizer.Default);
				formulas.Index = _worksheet.GetMaxShareFunctionIndex(isArray: false);
				formulas.IsArray = false;
				_worksheet._sharedFormulas.Add(formulas.Index, formulas);
			}
			else
			{
				flag = true;
			}
			formulas.StartCol = address._toCol + 1;
			if (address._fromRow < excelRange._fromRow)
			{
				formulas.StartRow = excelRange._fromRow;
			}
			else
			{
				formulas.StartRow = address._fromRow;
			}
			if (excelRange._toRow < address._toRow)
			{
				formulas.Address = ExcelCellBase.GetAddress(formulas.StartRow, formulas.StartCol, excelRange._toRow, excelRange._toCol);
			}
			else
			{
				formulas.Address = ExcelCellBase.GetAddress(formulas.StartRow, formulas.StartCol, address._toRow, excelRange._toCol);
			}
			formulas.Formula = ExcelCellBase.TranslateFromR1C1(formulaR1C, formulas.StartRow, formulas.StartCol);
			_worksheet.Cells[formulas.Address].SetSharedFormulaID(formulas.Index);
		}
		if (excelRange._toRow > address._toRow)
		{
			if (flag)
			{
				formulas = new ExcelWorksheet.Formulas(SourceCodeTokenizer.Default);
				formulas.Index = _worksheet.GetMaxShareFunctionIndex(isArray: false);
				formulas.IsArray = false;
				_worksheet._sharedFormulas.Add(formulas.Index, formulas);
			}
			formulas.StartCol = excelRange._fromCol;
			formulas.StartRow = _toRow + 1;
			formulas.Formula = ExcelCellBase.TranslateFromR1C1(formulaR1C, formulas.StartRow, formulas.StartCol);
			formulas.Address = ExcelCellBase.GetAddress(formulas.StartRow, formulas.StartCol, excelRange._toRow, excelRange._toCol);
			_worksheet.Cells[formulas.Address].SetSharedFormulaID(formulas.Index);
		}
	}

	private object ConvertData(ExcelTextFormat Format, string v, int col, bool isText)
	{
		if (isText && (Format.DataTypes == null || Format.DataTypes.Length < col))
		{
			if (!string.IsNullOrEmpty(v))
			{
				return v;
			}
			return null;
		}
		double result;
		DateTime result2;
		if (Format.DataTypes == null || Format.DataTypes.Length <= col || Format.DataTypes[col] == eDataTypes.Unknown)
		{
			string text = (v.EndsWith("%") ? v.Substring(0, v.Length - 1) : v);
			if (double.TryParse(text, NumberStyles.Any, Format.Culture, out result))
			{
				if (text == v)
				{
					return result;
				}
				return result / 100.0;
			}
			if (DateTime.TryParse(v, Format.Culture, DateTimeStyles.None, out result2))
			{
				return result2;
			}
			if (!string.IsNullOrEmpty(v))
			{
				return v;
			}
			return null;
		}
		switch (Format.DataTypes[col])
		{
		case eDataTypes.Number:
			if (double.TryParse(v, NumberStyles.Any, Format.Culture, out result))
			{
				return result;
			}
			return v;
		case eDataTypes.DateTime:
			if (DateTime.TryParse(v, Format.Culture, DateTimeStyles.None, out result2))
			{
				return result2;
			}
			return v;
		case eDataTypes.Percent:
			if (double.TryParse(v.EndsWith("%") ? v.Substring(0, v.Length - 1) : v, NumberStyles.Any, Format.Culture, out result))
			{
				return result / 100.0;
			}
			return v;
		case eDataTypes.String:
			return v;
		default:
			if (!string.IsNullOrEmpty(v))
			{
				return v;
			}
			return null;
		}
	}

	public ExcelRangeBase LoadFromDataReader(IDataReader Reader, bool PrintHeaders, string TableName, TableStyles TableStyle = TableStyles.None)
	{
		ExcelRangeBase excelRangeBase = LoadFromDataReader(Reader, PrintHeaders);
		int num = excelRangeBase.Rows - 1;
		if (num >= 0 && excelRangeBase.Columns > 0)
		{
			ExcelTable excelTable = _worksheet.Tables.Add(new ExcelAddressBase(_fromRow, _fromCol, _fromRow + ((num <= 0) ? 1 : num), _fromCol + excelRangeBase.Columns - 1), TableName);
			excelTable.ShowHeader = PrintHeaders;
			excelTable.TableStyle = TableStyle;
		}
		return excelRangeBase;
	}

	public ExcelRangeBase LoadFromDataReader(IDataReader Reader, bool PrintHeaders)
	{
		if (Reader == null)
		{
			throw new ArgumentNullException("Reader", "Reader can't be null");
		}
		int fieldCount = Reader.FieldCount;
		int fromCol = _fromCol;
		int num = _fromRow;
		if (PrintHeaders)
		{
			for (int i = 0; i < fieldCount; i++)
			{
				_worksheet.SetValueInner(num, fromCol++, Reader.GetName(i));
			}
			num++;
			fromCol = _fromCol;
		}
		while (Reader.Read())
		{
			for (int j = 0; j < fieldCount; j++)
			{
				_worksheet.SetValueInner(num, fromCol++, Reader.GetValue(j));
			}
			num++;
			fromCol = _fromCol;
		}
		return _worksheet.Cells[_fromRow, _fromCol, num - 1, _fromCol + fieldCount - 1];
	}

	public ExcelRangeBase LoadFromDataTable(DataTable Table, bool PrintHeaders, TableStyles TableStyle)
	{
		ExcelRangeBase result = LoadFromDataTable(Table, PrintHeaders);
		int num = ((Table.Rows.Count == 0) ? 1 : Table.Rows.Count) + (PrintHeaders ? 1 : 0);
		if (num >= 0 && Table.Columns.Count > 0)
		{
			ExcelTable excelTable = _worksheet.Tables.Add(new ExcelAddressBase(_fromRow, _fromCol, _fromRow + num - 1, _fromCol + Table.Columns.Count - 1), Table.TableName);
			excelTable.ShowHeader = PrintHeaders;
			excelTable.TableStyle = TableStyle;
		}
		return result;
	}

	public ExcelRangeBase LoadFromDataTable(DataTable Table, bool PrintHeaders)
	{
		if (Table == null)
		{
			throw new ArgumentNullException("Table can't be null");
		}
		if (Table.Rows.Count == 0 && !PrintHeaders)
		{
			return null;
		}
		List<object[]> list2 = new List<object[]>();
		if (PrintHeaders)
		{
			list2.Add((from DataColumn dc in Table.Columns
				select dc.Caption).ToArray());
		}
		foreach (DataRow row in Table.Rows)
		{
			list2.Add(row.ItemArray);
		}
		_worksheet._values.SetRangeValueSpecial(_fromRow, _fromCol, _fromRow + list2.Count - 1, _fromCol + Table.Columns.Count - 1, delegate(List<ExcelCoreValue> list, int index, int rowIx, int columnIx, object value)
		{
			rowIx -= _fromRow;
			columnIx -= _fromCol;
			object obj = ((List<object[]>)value)[rowIx][columnIx];
			if (obj != null && obj != DBNull.Value && !string.IsNullOrEmpty(obj.ToString()))
			{
				list[index] = new ExcelCoreValue
				{
					_value = obj,
					_styleId = list[index]._styleId
				};
			}
		}, list2);
		return _worksheet.Cells[_fromRow, _fromCol, _fromRow + list2.Count - 1, _fromCol + Table.Columns.Count - 1];
	}

	public ExcelRangeBase LoadFromArrays(IEnumerable<object[]> Data)
	{
		if (Data == null)
		{
			throw new ArgumentNullException("data");
		}
		List<object[]> list2 = new List<object[]>();
		int num = 0;
		foreach (object[] Datum in Data)
		{
			list2.Add(Datum);
			if (num < Datum.Length)
			{
				num = Datum.Length;
			}
		}
		if (list2.Count == 0)
		{
			return null;
		}
		_worksheet._values.SetRangeValueSpecial(_fromRow, _fromCol, _fromRow + list2.Count - 1, _fromCol + num - 1, delegate(List<ExcelCoreValue> list, int index, int rowIx, int columnIx, object value)
		{
			rowIx -= _fromRow;
			columnIx -= _fromCol;
			List<object[]> list3 = (List<object[]>)value;
			if (list3.Count > rowIx)
			{
				object[] array = list3[rowIx];
				if (array.Length > columnIx)
				{
					object obj = array[columnIx];
					if (obj != null && obj != DBNull.Value && !string.IsNullOrEmpty(obj.ToString()))
					{
						list[index] = new ExcelCoreValue
						{
							_value = obj,
							_styleId = list[index]._styleId
						};
					}
				}
			}
		}, list2);
		return _worksheet.Cells[_fromRow, _fromCol, _fromRow + list2.Count - 1, _fromCol + num - 1];
	}

	public ExcelRangeBase LoadFromCollection<T>(IEnumerable<T> Collection)
	{
		return LoadFromCollection(Collection, PrintHeaders: false, TableStyles.None, BindingFlags.Instance | BindingFlags.Public, null);
	}

	public ExcelRangeBase LoadFromCollection<T>(IEnumerable<T> Collection, bool PrintHeaders)
	{
		return LoadFromCollection(Collection, PrintHeaders, TableStyles.None, BindingFlags.Instance | BindingFlags.Public, null);
	}

	public ExcelRangeBase LoadFromCollection<T>(IEnumerable<T> Collection, bool PrintHeaders, TableStyles TableStyle)
	{
		return LoadFromCollection(Collection, PrintHeaders, TableStyle, BindingFlags.Instance | BindingFlags.Public, null);
	}

	public ExcelRangeBase LoadFromCollection<T>(IEnumerable<T> Collection, bool PrintHeaders, TableStyles TableStyle, BindingFlags memberFlags, MemberInfo[] Members)
	{
		Type typeFromHandle = typeof(T);
		bool flag = true;
		if (Members == null)
		{
			Members = typeFromHandle.GetProperties(memberFlags);
		}
		else
		{
			if (Members.Length == 0)
			{
				throw new ArgumentException("Parameter Members must have at least one property. Length is zero");
			}
			MemberInfo[] array = Members;
			foreach (MemberInfo memberInfo in array)
			{
				if (memberInfo.DeclaringType != null && memberInfo.DeclaringType != typeFromHandle)
				{
					flag = false;
				}
				if (memberInfo.DeclaringType != null && memberInfo.DeclaringType != typeFromHandle && !TypeCompat.IsSubclassOf(typeFromHandle, memberInfo.DeclaringType) && !TypeCompat.IsSubclassOf(memberInfo.DeclaringType, typeFromHandle))
				{
					throw new InvalidCastException("Supplied\u00a0properties\u00a0in\u00a0parameter\u00a0Properties\u00a0must\u00a0be\u00a0of\u00a0the\u00a0same\u00a0type\u00a0as\u00a0T\u00a0(or\u00a0an\u00a0assignable\u00a0type\u00a0from\u00a0T)");
				}
			}
		}
		object[,] array2 = new object[PrintHeaders ? (Collection.Count() + 1) : Collection.Count(), Members.Count()];
		int num = 0;
		int num2 = 0;
		if (Members.Length != 0 && PrintHeaders)
		{
			MemberInfo[] array = Members;
			foreach (MemberInfo memberInfo2 in array)
			{
				DescriptionAttribute descriptionAttribute = memberInfo2.GetCustomAttributes(typeof(DescriptionAttribute), inherit: false).FirstOrDefault() as DescriptionAttribute;
				string empty = string.Empty;
				empty = ((descriptionAttribute == null) ? ((!(memberInfo2.GetCustomAttributes(typeof(DisplayNameAttribute), inherit: false).FirstOrDefault() is DisplayNameAttribute displayNameAttribute)) ? memberInfo2.Name.Replace('_', ' ') : displayNameAttribute.DisplayName) : descriptionAttribute.Description);
				array2[num2, num++] = empty;
			}
			num2++;
		}
		if (!Collection.Any() && (Members.Length == 0 || !PrintHeaders))
		{
			return null;
		}
		foreach (T item in Collection)
		{
			num = 0;
			if (item is string || item is decimal || item is DateTime || TypeCompat.IsPrimitive(item))
			{
				array2[num2, num++] = item;
			}
			else
			{
				MemberInfo[] array = Members;
				foreach (MemberInfo memberInfo3 in array)
				{
					if (!flag && item.GetType().GetMember(memberInfo3.Name, memberFlags).Length == 0)
					{
						num++;
					}
					else if (memberInfo3 is PropertyInfo)
					{
						array2[num2, num++] = ((PropertyInfo)memberInfo3).GetValue(item, null);
					}
					else if (memberInfo3 is FieldInfo)
					{
						array2[num2, num++] = ((FieldInfo)memberInfo3).GetValue(item);
					}
					else if (memberInfo3 is MethodInfo)
					{
						array2[num2, num++] = ((MethodInfo)memberInfo3).Invoke(item, null);
					}
				}
			}
			num2++;
		}
		_worksheet.SetRangeValueInner(_fromRow, _fromCol, _fromRow + num2 - 1, _fromCol + num - 1, array2);
		if (num2 == 1 && PrintHeaders)
		{
			num2++;
		}
		ExcelRange excelRange = _worksheet.Cells[_fromRow, _fromCol, _fromRow + num2 - 1, _fromCol + num - 1];
		if (TableStyle != 0)
		{
			ExcelTable excelTable = _worksheet.Tables.Add(excelRange, "");
			excelTable.ShowHeader = PrintHeaders;
			excelTable.TableStyle = TableStyle;
		}
		return excelRange;
	}

	public ExcelRangeBase LoadFromText(string Text)
	{
		return LoadFromText(Text, new ExcelTextFormat());
	}

	public ExcelRangeBase LoadFromText(string Text, ExcelTextFormat Format)
	{
		if (string.IsNullOrEmpty(Text))
		{
			ExcelRange excelRange = _worksheet.Cells[_fromRow, _fromCol];
			excelRange.Value = "";
			return excelRange;
		}
		if (Format == null)
		{
			Format = new ExcelTextFormat();
		}
		string[] array = ((Format.TextQualifier != 0) ? GetLines(Text, Format) : Regex.Split(Text, Format.EOL));
		int num = 0;
		int num2 = 0;
		int num3 = num2;
		int num4 = 1;
		List<object>[] values = new List<object>[array.Length];
		string[] array2 = array;
		foreach (string text in array2)
		{
			List<object> list2 = new List<object>();
			values[num] = list2;
			if (num4 > Format.SkipLinesBeginning && num4 <= array.Length - Format.SkipLinesEnd)
			{
				num2 = 0;
				string text2 = "";
				bool flag = false;
				bool flag2 = false;
				int num5 = 0;
				int num6 = 0;
				string text3 = text;
				for (int j = 0; j < text3.Length; j++)
				{
					char c = text3[j];
					if (Format.TextQualifier != 0 && c == Format.TextQualifier)
					{
						if (!flag && text2 != "")
						{
							throw new Exception($"Invalid Text Qualifier in line : {text}");
						}
						flag2 = !flag2;
						num5++;
						num6++;
						flag = true;
						continue;
					}
					if (num5 > 1 && !string.IsNullOrEmpty(text2))
					{
						text2 += new string(Format.TextQualifier, num5 / 2);
					}
					else if (num5 > 2 && string.IsNullOrEmpty(text2))
					{
						text2 += new string(Format.TextQualifier, (num5 - 1) / 2);
					}
					if (flag2)
					{
						text2 += c;
					}
					else if (c == Format.Delimiter)
					{
						list2.Add(ConvertData(Format, text2, num2, flag));
						text2 = "";
						flag = false;
						num2++;
					}
					else
					{
						if (num5 % 2 == 1)
						{
							throw new Exception($"Text delimiter is not closed in line : {text}");
						}
						text2 += c;
					}
					num5 = 0;
				}
				if (num5 > 1 && text2 != "" && num5 == 2)
				{
					text2 += new string(Format.TextQualifier, num5 / 2);
				}
				if (num6 % 2 == 1)
				{
					throw new Exception($"Text delimiter is not closed in line : {text}");
				}
				list2.Add(ConvertData(Format, text2, num2, flag));
				if (num2 > num3)
				{
					num3 = num2;
				}
				num++;
			}
			num4++;
		}
		_worksheet._values.SetRangeValueSpecial(_fromRow, _fromCol, _fromRow + values.Length - 1, _fromCol + num3, delegate(List<ExcelCoreValue> list, int index, int rowIx, int columnIx, object value)
		{
			rowIx -= _fromRow;
			columnIx -= _fromCol;
			List<object> list3 = values[rowIx];
			if (list3 != null && list3.Count > columnIx)
			{
				list[index] = new ExcelCoreValue
				{
					_value = list3[columnIx],
					_styleId = list[index]._styleId
				};
			}
		}, values);
		return _worksheet.Cells[_fromRow, _fromCol, _fromRow + num - 1, _fromCol + num3];
	}

	private string[] GetLines(string text, ExcelTextFormat Format)
	{
		if (Format.EOL == null || Format.EOL.Length == 0)
		{
			return new string[1] { text };
		}
		string eOL = Format.EOL;
		List<string> list = new List<string>();
		bool flag = false;
		int num = 0;
		for (int i = 0; i < text.Length; i++)
		{
			if (text[i] == Format.TextQualifier)
			{
				flag = !flag;
			}
			else if (!flag && IsEOL(text, i, eOL))
			{
				list.Add(text.Substring(num, i - num));
				i += eOL.Length - 1;
				num = i + 1;
			}
		}
		if (flag)
		{
			throw new ArgumentException($"Text delimiter is not closed in line : {list.Count}");
		}
		if (num >= Format.EOL.Length && IsEOL(text, num - Format.EOL.Length, Format.EOL))
		{
			list.Add("");
		}
		else
		{
			list.Add(text.Substring(num));
		}
		return list.ToArray();
	}

	private bool IsEOL(string text, int ix, string eol)
	{
		for (int i = 0; i < eol.Length; i++)
		{
			if (text[ix + i] != eol[i])
			{
				return false;
			}
		}
		return ix + eol.Length <= text.Length;
	}

	public ExcelRangeBase LoadFromText(string Text, ExcelTextFormat Format, TableStyles TableStyle, bool FirstRowIsHeader)
	{
		ExcelRangeBase excelRangeBase = LoadFromText(Text, Format);
		ExcelTable excelTable = _worksheet.Tables.Add(excelRangeBase, "");
		excelTable.ShowHeader = FirstRowIsHeader;
		excelTable.TableStyle = TableStyle;
		return excelRangeBase;
	}

	public ExcelRangeBase LoadFromText(FileInfo TextFile)
	{
		return LoadFromText(File.ReadAllText(TextFile.FullName, Encoding.ASCII));
	}

	public ExcelRangeBase LoadFromText(FileInfo TextFile, ExcelTextFormat Format)
	{
		return LoadFromText(File.ReadAllText(TextFile.FullName, Format.Encoding), Format);
	}

	public ExcelRangeBase LoadFromText(FileInfo TextFile, ExcelTextFormat Format, TableStyles TableStyle, bool FirstRowIsHeader)
	{
		return LoadFromText(File.ReadAllText(TextFile.FullName, Format.Encoding), Format, TableStyle, FirstRowIsHeader);
	}

	public T GetValue<T>()
	{
		return ConvertUtil.GetTypedCellValue<T>(Value);
	}

	public ExcelRangeBase Offset(int RowOffset, int ColumnOffset)
	{
		if (_fromRow + RowOffset < 1 || _fromCol + ColumnOffset < 1 || _fromRow + RowOffset > 1048576 || _fromCol + ColumnOffset > 16384)
		{
			throw new ArgumentOutOfRangeException("Offset value out of range");
		}
		string address = ExcelCellBase.GetAddress(_fromRow + RowOffset, _fromCol + ColumnOffset, _toRow + RowOffset, _toCol + ColumnOffset);
		return new ExcelRangeBase(_worksheet, address);
	}

	public ExcelRangeBase Offset(int RowOffset, int ColumnOffset, int NumberOfRows, int NumberOfColumns)
	{
		if (NumberOfRows < 1 || NumberOfColumns < 1)
		{
			throw new Exception("Number of rows/columns must be greater than 0");
		}
		NumberOfRows--;
		NumberOfColumns--;
		if (_fromRow + RowOffset < 1 || _fromCol + ColumnOffset < 1 || _fromRow + RowOffset > 1048576 || _fromCol + ColumnOffset > 16384 || _fromRow + RowOffset + NumberOfRows < 1 || _fromCol + ColumnOffset + NumberOfColumns < 1 || _fromRow + RowOffset + NumberOfRows > 1048576 || _fromCol + ColumnOffset + NumberOfColumns > 16384)
		{
			throw new ArgumentOutOfRangeException("Offset value out of range");
		}
		string address = ExcelCellBase.GetAddress(_fromRow + RowOffset, _fromCol + ColumnOffset, _fromRow + RowOffset + NumberOfRows, _fromCol + ColumnOffset + NumberOfColumns);
		return new ExcelRangeBase(_worksheet, address);
	}

	public ExcelComment AddComment(string Text, string Author)
	{
		if (string.IsNullOrEmpty(Author))
		{
			Author = Thread.CurrentPrincipal.Identity.Name;
		}
		_changePropMethod(this, _setExistsCommentDelegate, null);
		_changePropMethod(this, _setCommentDelegate, new string[2] { Text, Author });
		return _worksheet.Comments[new ExcelCellAddress(_fromRow, _fromCol)];
	}

	public void Copy(ExcelRangeBase Destination)
	{
		Copy(Destination, null);
	}

	public void Copy(ExcelRangeBase Destination, ExcelRangeCopyOptionFlags? excelRangeCopyOptionFlags)
	{
		bool flag = Destination._worksheet.Workbook == _worksheet.Workbook;
		ExcelStyles styles = _worksheet.Workbook.Styles;
		ExcelStyles styles2 = Destination._worksheet.Workbook.Styles;
		Dictionary<int, int> dictionary = new Dictionary<int, int>();
		int num = _toRow - _fromRow + 1;
		int num2 = _toCol - _fromCol + 1;
		int styleId = 0;
		object value = null;
		byte value2 = 0;
		Uri value3 = null;
		bool flag2 = excelRangeCopyOptionFlags.HasValue && (excelRangeCopyOptionFlags.Value & ExcelRangeCopyOptionFlags.ExcludeFormulas) == ExcelRangeCopyOptionFlags.ExcludeFormulas;
		CellsStoreEnumerator<ExcelCoreValue> cellsStoreEnumerator = new CellsStoreEnumerator<ExcelCoreValue>(_worksheet._values, _fromRow, _fromCol, _toRow, _toCol);
		List<CopiedCell> list = new List<CopiedCell>();
		while (cellsStoreEnumerator.Next())
		{
			int row = cellsStoreEnumerator.Row;
			int column = cellsStoreEnumerator.Column;
			CopiedCell copiedCell = new CopiedCell
			{
				Row = Destination._fromRow + (row - _fromRow),
				Column = Destination._fromCol + (column - _fromCol),
				Value = cellsStoreEnumerator.Value._value
			};
			if (!flag2 && _worksheet._formulas.Exists(row, column, ref value))
			{
				if (value is int)
				{
					copiedCell.Formula = _worksheet.GetFormula(cellsStoreEnumerator.Row, cellsStoreEnumerator.Column);
					if (_worksheet._flags.GetFlagValue(cellsStoreEnumerator.Row, cellsStoreEnumerator.Column, CellFlags.ArrayFormula))
					{
						Destination._worksheet._flags.SetFlagValue(cellsStoreEnumerator.Row, cellsStoreEnumerator.Column, value: true, CellFlags.ArrayFormula);
					}
				}
				else
				{
					copiedCell.Formula = value;
				}
			}
			if (_worksheet.ExistsStyleInner(row, column, ref styleId))
			{
				if (flag)
				{
					copiedCell.StyleID = styleId;
				}
				else
				{
					if (dictionary.ContainsKey(styleId))
					{
						styleId = dictionary[styleId];
					}
					else
					{
						int key = styleId;
						styleId = styles2.CloneStyle(styles, styleId);
						dictionary.Add(key, styleId);
					}
					copiedCell.StyleID = styleId;
				}
			}
			if (_worksheet._hyperLinks.Exists(row, column, ref value3))
			{
				copiedCell.HyperLink = value3;
			}
			copiedCell.Comment = _worksheet.Cells[cellsStoreEnumerator.Row, cellsStoreEnumerator.Column].Comment;
			if (_worksheet._flags.Exists(row, column, ref value2))
			{
				copiedCell.Flag = value2;
			}
			list.Add(copiedCell);
		}
		CellsStoreEnumerator<ExcelCoreValue> cellsStoreEnumerator2 = new CellsStoreEnumerator<ExcelCoreValue>(_worksheet._values, _fromRow, _fromCol, _toRow, _toCol);
		while (cellsStoreEnumerator2.Next())
		{
			if (_worksheet.ExistsValueInner(cellsStoreEnumerator2.Row, cellsStoreEnumerator2.Column))
			{
				continue;
			}
			int row2 = Destination._fromRow + (cellsStoreEnumerator2.Row - _fromRow);
			int column2 = Destination._fromCol + (cellsStoreEnumerator2.Column - _fromCol);
			CopiedCell copiedCell2 = new CopiedCell
			{
				Row = row2,
				Column = column2,
				Value = null
			};
			styleId = cellsStoreEnumerator2.Value._styleId;
			if (flag)
			{
				copiedCell2.StyleID = styleId;
			}
			else
			{
				if (dictionary.ContainsKey(styleId))
				{
					styleId = dictionary[styleId];
				}
				else
				{
					int key2 = styleId;
					styleId = styles2.CloneStyle(styles, styleId);
					dictionary.Add(key2, styleId);
				}
				copiedCell2.StyleID = styleId;
			}
			list.Add(copiedCell2);
		}
		Dictionary<int, ExcelAddress> dictionary2 = new Dictionary<int, ExcelAddress>();
		CellsStoreEnumerator<int> cellsStoreEnumerator3 = new CellsStoreEnumerator<int>(_worksheet.MergedCells._cells, _fromRow, _fromCol, _toRow, _toCol);
		while (cellsStoreEnumerator3.Next())
		{
			if (!dictionary2.ContainsKey(cellsStoreEnumerator3.Value))
			{
				ExcelAddress excelAddress = new ExcelAddress(_worksheet.Name, _worksheet.MergedCells.List[cellsStoreEnumerator3.Value]);
				if (Collide(excelAddress) == eAddressCollition.Inside)
				{
					dictionary2.Add(cellsStoreEnumerator3.Value, new ExcelAddress(Destination._fromRow + (excelAddress.Start.Row - _fromRow), Destination._fromCol + (excelAddress.Start.Column - _fromCol), Destination._fromRow + (excelAddress.End.Row - _fromRow), Destination._fromCol + (excelAddress.End.Column - _fromCol)));
				}
				else
				{
					dictionary2.Add(cellsStoreEnumerator3.Value, null);
				}
			}
		}
		Destination._worksheet.MergedCells.Clear(new ExcelAddressBase(Destination._fromRow, Destination._fromCol, Destination._fromRow + num - 1, Destination._fromCol + num2 - 1));
		Destination._worksheet._values.Clear(Destination._fromRow, Destination._fromCol, num, num2);
		Destination._worksheet._formulas.Clear(Destination._fromRow, Destination._fromCol, num, num2);
		Destination._worksheet._hyperLinks.Clear(Destination._fromRow, Destination._fromCol, num, num2);
		Destination._worksheet._flags.Clear(Destination._fromRow, Destination._fromCol, num, num2);
		Destination._worksheet._commentsStore.Clear(Destination._fromRow, Destination._fromCol, num, num2);
		foreach (CopiedCell item in list)
		{
			Destination._worksheet.SetValueInner(item.Row, item.Column, item.Value);
			if (item.StyleID.HasValue)
			{
				Destination._worksheet.SetStyleInner(item.Row, item.Column, item.StyleID.Value);
			}
			if (item.Formula != null)
			{
				item.Formula = ExcelCellBase.UpdateFormulaReferences(item.Formula.ToString(), Destination._fromRow - _fromRow, Destination._fromCol - _fromCol, 0, 0, Destination.WorkSheet, Destination.WorkSheet, setFixed: true);
				Destination._worksheet._formulas.SetValue(item.Row, item.Column, item.Formula);
			}
			if (item.HyperLink != null)
			{
				Destination._worksheet._hyperLinks.SetValue(item.Row, item.Column, item.HyperLink);
			}
			if (item.Comment != null)
			{
				Destination.Worksheet.Cells[item.Row, item.Column].AddComment(item.Comment.Text, item.Comment.Author);
			}
			if (item.Flag != 0)
			{
				Destination._worksheet._flags.SetValue(item.Row, item.Column, item.Flag);
			}
		}
		foreach (ExcelAddress value4 in dictionary2.Values)
		{
			if (value4 != null)
			{
				Destination._worksheet.MergedCells.Add(value4, doValidate: true);
			}
		}
		if (_fromCol == 1 && _toCol == 16384)
		{
			for (int i = 0; i < base.Rows; i++)
			{
				Destination.Worksheet.Row(Destination.Start.Row + i).OutlineLevel = Worksheet.Row(_fromRow + i).OutlineLevel;
			}
		}
		if (_fromRow == 1 && _toRow == 1048576)
		{
			for (int j = 0; j < base.Columns; j++)
			{
				Destination.Worksheet.Column(Destination.Start.Column + j).OutlineLevel = Worksheet.Column(_fromCol + j).OutlineLevel;
			}
		}
	}

	public void Clear()
	{
		Delete(this, shift: false);
	}

	public void CreateArrayFormula(string ArrayFormula)
	{
		if (Addresses != null)
		{
			throw new Exception("An Arrayformula can not have more than one address");
		}
		Set_SharedFormula(this, ArrayFormula, this, IsArray: true);
	}

	internal void Delete(ExcelAddressBase Range, bool shift)
	{
		_worksheet.MergedCells.Clear(Range);
		ExcelAddressBase dimension = Worksheet.Dimension;
		int num = ((dimension == null || Range._fromRow > dimension._fromRow || Range._toRow < dimension._toRow) ? Range._fromRow : 0);
		int num2 = ((dimension == null || Range._fromCol > dimension._fromCol || Range._toCol < dimension._toCol) ? Range._fromCol : 0);
		int rows = Range._toRow - num + 1;
		int columns = Range._toCol - num2 + 1;
		_worksheet._values.Delete(num, num2, rows, columns, shift);
		_worksheet._formulas.Delete(num, num2, rows, columns, shift);
		_worksheet._hyperLinks.Delete(num, num2, rows, columns, shift);
		_worksheet._flags.Delete(num, num2, rows, columns, shift);
		_worksheet._commentsStore.Delete(num, num2, rows, columns, shift);
		if (Addresses == null)
		{
			return;
		}
		foreach (ExcelAddress address in Addresses)
		{
			Delete(address, shift);
		}
	}

	private void DeleteCheckMergedCells(ExcelAddressBase Range)
	{
		List<string> list = new List<string>();
		foreach (string mergedCell in Worksheet.MergedCells)
		{
			switch (Range.Collide(new ExcelAddress(Range.WorkSheet, mergedCell)))
			{
			case eAddressCollition.Inside:
				list.Add(mergedCell);
				break;
			default:
				throw new InvalidOperationException("Can't remove/overwrite a part of cells that are merged");
			case eAddressCollition.No:
				break;
			}
		}
		foreach (string item in list)
		{
			Worksheet.MergedCells.Remove(item);
		}
	}

	public void Dispose()
	{
	}

	public IEnumerator<ExcelRangeBase> GetEnumerator()
	{
		Reset();
		return this;
	}

	IEnumerator IEnumerable.GetEnumerator()
	{
		Reset();
		return this;
	}

	public bool MoveNext()
	{
		if (cellEnum.Next())
		{
			return true;
		}
		if (_addresses != null)
		{
			_enumAddressIx++;
			if (_enumAddressIx < _addresses.Count)
			{
				cellEnum = new CellsStoreEnumerator<ExcelCoreValue>(_worksheet._values, _addresses[_enumAddressIx]._fromRow, _addresses[_enumAddressIx]._fromCol, _addresses[_enumAddressIx]._toRow, _addresses[_enumAddressIx]._toCol);
				return MoveNext();
			}
			return false;
		}
		return false;
	}

	public void Reset()
	{
		_enumAddressIx = -1;
		cellEnum = new CellsStoreEnumerator<ExcelCoreValue>(_worksheet._values, _fromRow, _fromCol, _toRow, _toCol);
	}

	public void Sort()
	{
		Sort(new int[1], new bool[1]);
	}

	public void Sort(int column, bool descending = false)
	{
		Sort(new int[1] { column }, new bool[1] { descending });
	}

	public void Sort(int[] columns, bool[] descending = null, CultureInfo culture = null, CompareOptions compareOptions = CompareOptions.None)
	{
		if (columns == null)
		{
			columns = new int[1];
		}
		int num = _toCol - _fromCol + 1;
		int[] array = columns;
		foreach (int num2 in array)
		{
			if (num2 > num - 1 || num2 < 0)
			{
				throw new ArgumentException("Can not reference columns outside the boundries of the range. Note that column reference is zero-based within the range");
			}
		}
		CellsStoreEnumerator<ExcelCoreValue> cellsStoreEnumerator = new CellsStoreEnumerator<ExcelCoreValue>(_worksheet._values, _fromRow, _fromCol, _toRow, _toCol);
		List<SortItem<ExcelCoreValue>> list = new List<SortItem<ExcelCoreValue>>();
		SortItem<ExcelCoreValue> item = default(SortItem<ExcelCoreValue>);
		while (cellsStoreEnumerator.Next())
		{
			if (list.Count == 0 || list[list.Count - 1].Row != cellsStoreEnumerator.Row)
			{
				SortItem<ExcelCoreValue> sortItem = default(SortItem<ExcelCoreValue>);
				sortItem.Row = cellsStoreEnumerator.Row;
				sortItem.Items = new ExcelCoreValue[num];
				item = sortItem;
				list.Add(item);
			}
			item.Items[cellsStoreEnumerator.Column - _fromCol] = cellsStoreEnumerator.Value;
		}
		if (descending == null)
		{
			descending = new bool[columns.Length];
			for (int j = 0; j < columns.Length; j++)
			{
				descending[j] = false;
			}
		}
		Comp comp = new Comp();
		comp.columns = columns;
		comp.descending = descending;
		comp.cultureInfo = culture ?? CultureInfo.CurrentCulture;
		comp.compareOptions = compareOptions;
		list.Sort(comp);
		Dictionary<string, byte> items = GetItems(_worksheet._flags, _fromRow, _fromCol, _toRow, _toCol);
		Dictionary<string, object> items2 = GetItems(_worksheet._formulas, _fromRow, _fromCol, _toRow, _toCol);
		Dictionary<string, Uri> items3 = GetItems(_worksheet._hyperLinks, _fromRow, _fromCol, _toRow, _toCol);
		Dictionary<string, int> items4 = GetItems(_worksheet._commentsStore, _fromRow, _fromCol, _toRow, _toCol);
		HashSet<int> hashSet = new HashSet<int>();
		_worksheet._values.Clear(_fromRow, _fromCol, _toRow - _fromRow + 1, num);
		for (int k = 0; k < list.Count; k++)
		{
			for (int l = 0; l < num; l++)
			{
				int num3 = _fromRow + k;
				int num4 = _fromCol + l;
				_worksheet._values.SetValueSpecial(num3, num4, SortSetValue, list[k].Items[l]);
				string address = ExcelCellBase.GetAddress(list[k].Row, _fromCol + l);
				if (items.ContainsKey(address))
				{
					_worksheet._flags.SetValue(num3, num4, items[address]);
				}
				if (items2.ContainsKey(address))
				{
					_worksheet._formulas.SetValue(num3, num4, items2[address]);
					if (items2[address] is int)
					{
						int num5 = (int)items2[address];
						if (!hashSet.Contains(num5))
						{
							ExcelAddress excelAddress = new ExcelAddress(Worksheet._sharedFormulas[num5].Address);
							if (excelAddress._fromRow > num3)
							{
								ExcelWorksheet.Formulas formulas = Worksheet._sharedFormulas[num5];
								formulas.Formula = ExcelCellBase.TranslateFromR1C1(ExcelCellBase.TranslateToR1C1(formulas.Formula, formulas.StartRow, formulas.StartCol), num3, formulas.StartCol);
								formulas.StartRow = num3;
								formulas.Address = ExcelCellBase.GetAddress(num3, num4, excelAddress._toRow, excelAddress._toCol);
							}
						}
						hashSet.Add(num5);
					}
				}
				if (items3.ContainsKey(address))
				{
					_worksheet._hyperLinks.SetValue(num3, num4, items3[address]);
				}
				if (items4.ContainsKey(address))
				{
					int num6 = items4[address];
					_worksheet._commentsStore.SetValue(num3, num4, num6);
					_worksheet._comments[num6].Address = ExcelCellBase.GetAddress(num3, num4);
				}
			}
		}
	}

	private static Dictionary<string, T> GetItems<T>(CellStore<T> store, int fromRow, int fromCol, int toRow, int toCol)
	{
		CellsStoreEnumerator<T> cellsStoreEnumerator = new CellsStoreEnumerator<T>(store, fromRow, fromCol, toRow, toCol);
		Dictionary<string, T> dictionary = new Dictionary<string, T>();
		while (cellsStoreEnumerator.Next())
		{
			dictionary.Add(cellsStoreEnumerator.CellAddress, cellsStoreEnumerator.Value);
		}
		return dictionary;
	}

	private static void SortSetValue(List<ExcelCoreValue> list, int index, object value)
	{
		ExcelCoreValue excelCoreValue = (ExcelCoreValue)value;
		list[index] = new ExcelCoreValue
		{
			_value = excelCoreValue._value,
			_styleId = excelCoreValue._styleId
		};
	}
}
