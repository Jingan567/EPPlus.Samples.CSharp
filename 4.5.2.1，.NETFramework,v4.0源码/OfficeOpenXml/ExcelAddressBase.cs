using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml;

public class ExcelAddressBase : ExcelCellBase
{
	internal enum eAddressCollition
	{
		No,
		Partly,
		Inside,
		Equal
	}

	internal enum eShiftType
	{
		Right,
		Down,
		EntireRow,
		EntireColumn
	}

	internal enum AddressType
	{
		Invalid,
		InternalAddress,
		ExternalAddress,
		InternalName,
		ExternalName,
		Formula,
		R1C1
	}

	protected internal int _fromRow = -1;

	protected internal int _toRow;

	protected internal int _fromCol;

	protected internal int _toCol;

	protected internal bool _fromRowFixed;

	protected internal bool _fromColFixed;

	protected internal bool _toRowFixed;

	protected internal bool _toColFixed;

	protected internal string _wb;

	protected internal string _ws;

	protected internal string _address;

	protected ExcelCellAddress _start;

	protected ExcelCellAddress _end;

	protected ExcelTableAddress _table;

	private string _firstAddress;

	protected internal List<ExcelAddress> _addresses;

	private static readonly HashSet<char> FormulaCharacters = new HashSet<char>(new char[11]
	{
		'(', ')', '+', '-', '*', '/', '=', '^', '&', '%',
		'"'
	});

	public ExcelCellAddress Start
	{
		get
		{
			if (_start == null)
			{
				_start = new ExcelCellAddress(_fromRow, _fromCol);
			}
			return _start;
		}
	}

	public ExcelCellAddress End
	{
		get
		{
			if (_end == null)
			{
				_end = new ExcelCellAddress(_toRow, _toCol);
			}
			return _end;
		}
	}

	public ExcelTableAddress Table => _table;

	public virtual string Address => _address;

	internal string FullAddress
	{
		get
		{
			string text = "";
			if (_addresses != null)
			{
				foreach (ExcelAddress address in _addresses)
				{
					text = text + "," + address.GetAddress();
				}
			}
			else
			{
				text = GetAddress();
			}
			return text;
		}
	}

	public bool IsName => _fromRow < 0;

	internal string FirstAddress
	{
		get
		{
			if (string.IsNullOrEmpty(_firstAddress))
			{
				return _address;
			}
			return _firstAddress;
		}
	}

	internal string AddressSpaceSeparated => _address.Replace(',', ' ');

	internal string WorkSheet => _ws;

	internal virtual List<ExcelAddress> Addresses => _addresses;

	public int Rows => _toRow - _fromRow + 1;

	public int Columns => _toCol - _fromCol + 1;

	internal bool IsSingleCell
	{
		get
		{
			if (_fromRow == _toRow)
			{
				return _fromCol == _toCol;
			}
			return false;
		}
	}

	internal ExcelAddressBase()
	{
	}

	public ExcelAddressBase(int fromRow, int fromCol, int toRow, int toColumn)
	{
		_fromRow = fromRow;
		_toRow = toRow;
		_fromCol = fromCol;
		_toCol = toColumn;
		Validate();
		_address = ExcelCellBase.GetAddress(_fromRow, _fromCol, _toRow, _toCol);
	}

	public ExcelAddressBase(string worksheetName, int fromRow, int fromCol, int toRow, int toColumn)
	{
		_ws = worksheetName;
		_fromRow = fromRow;
		_toRow = toRow;
		_fromCol = fromCol;
		_toCol = toColumn;
		Validate();
		_address = ExcelCellBase.GetAddress(_fromRow, _fromCol, _toRow, _toCol);
	}

	public ExcelAddressBase(int fromRow, int fromCol, int toRow, int toColumn, bool fromRowFixed, bool fromColFixed, bool toRowFixed, bool toColFixed)
	{
		_fromRow = fromRow;
		_toRow = toRow;
		_fromCol = fromCol;
		_toCol = toColumn;
		_fromRowFixed = fromRowFixed;
		_fromColFixed = fromColFixed;
		_toRowFixed = toRowFixed;
		_toColFixed = toColFixed;
		Validate();
		_address = ExcelCellBase.GetAddress(_fromRow, _fromCol, _toRow, _toCol, _fromRowFixed, fromColFixed, _toRowFixed, _toColFixed);
	}

	public ExcelAddressBase(string address)
	{
		SetAddress(address);
	}

	public ExcelAddressBase(string address, ExcelPackage pck, ExcelAddressBase referenceAddress)
	{
		SetAddress(address);
		SetRCFromTable(pck, referenceAddress);
	}

	internal void SetRCFromTable(ExcelPackage pck, ExcelAddressBase referenceAddress)
	{
		if (!string.IsNullOrEmpty(_wb) || Table == null)
		{
			return;
		}
		foreach (ExcelWorksheet worksheet in pck.Workbook.Worksheets)
		{
			foreach (ExcelTable table in worksheet.Tables)
			{
				if (!table.Name.Equals(Table.Name, StringComparison.OrdinalIgnoreCase))
				{
					continue;
				}
				_ws = worksheet.Name;
				if (Table.IsAll)
				{
					_fromRow = table.Address._fromRow;
					_toRow = table.Address._toRow;
				}
				else if (Table.IsThisRow)
				{
					if (referenceAddress == null)
					{
						_fromRow = -1;
						_toRow = -1;
					}
					else
					{
						_fromRow = referenceAddress._fromRow;
						_toRow = _fromRow;
					}
				}
				else if (Table.IsHeader && Table.IsData)
				{
					_fromRow = table.Address._fromRow;
					_toRow = (table.ShowTotal ? (table.Address._toRow - 1) : table.Address._toRow);
				}
				else if (Table.IsData && Table.IsTotals)
				{
					_fromRow = (table.ShowHeader ? (table.Address._fromRow + 1) : table.Address._fromRow);
					_toRow = table.Address._toRow;
				}
				else if (Table.IsHeader)
				{
					_fromRow = (table.ShowHeader ? table.Address._fromRow : (-1));
					_toRow = (table.ShowHeader ? table.Address._fromRow : (-1));
				}
				else if (Table.IsTotals)
				{
					_fromRow = (table.ShowTotal ? table.Address._toRow : (-1));
					_toRow = (table.ShowTotal ? table.Address._toRow : (-1));
				}
				else
				{
					_fromRow = (table.ShowHeader ? (table.Address._fromRow + 1) : table.Address._fromRow);
					_toRow = (table.ShowTotal ? (table.Address._toRow - 1) : table.Address._toRow);
				}
				if (string.IsNullOrEmpty(Table.ColumnSpan))
				{
					_fromCol = table.Address._fromCol;
					_toCol = table.Address._toCol;
					return;
				}
				int num = table.Address._fromCol;
				string[] array = Table.ColumnSpan.Split(':');
				foreach (ExcelTableColumn item in (IEnumerable<ExcelTableColumn>)table.Columns)
				{
					if (_fromCol <= 0 && array[0].Equals(item.Name, StringComparison.OrdinalIgnoreCase))
					{
						_fromCol = num;
						if (array.Length == 1)
						{
							_toCol = _fromCol;
							return;
						}
					}
					else if (array.Length > 1 && _fromCol > 0 && array[1].Equals(item.Name, StringComparison.OrdinalIgnoreCase))
					{
						_toCol = num;
						return;
					}
					num++;
				}
			}
		}
	}

	internal ExcelAddressBase(string address, bool isName)
	{
		if (isName)
		{
			_address = address;
			_fromRow = -1;
			_fromCol = -1;
			_toRow = -1;
			_toCol = -1;
			_start = null;
			_end = null;
		}
		else
		{
			SetAddress(address);
		}
	}

	protected internal void SetAddress(string address)
	{
		address = address.Trim();
		if (ConvertUtil._invariantCompareInfo.IsPrefix(address, "'") || ConvertUtil._invariantCompareInfo.IsPrefix(address, "["))
		{
			SetWbWs(address);
		}
		else
		{
			_address = address;
		}
		_addresses = null;
		if (_address.IndexOfAny(new char[3] { ',', '!', '[' }) > -1)
		{
			ExtractAddress(_address);
		}
		else
		{
			ExcelCellBase.GetRowColFromAddress(_address, out _fromRow, out _fromCol, out _toRow, out _toCol, out _fromRowFixed, out _fromColFixed, out _toRowFixed, out _toColFixed);
			_start = null;
			_end = null;
		}
		_address = address;
		Validate();
	}

	protected internal virtual void ChangeAddress()
	{
	}

	private void SetWbWs(string address)
	{
		int num;
		if (address[0] == '[')
		{
			num = address.IndexOf("]");
			_wb = address.Substring(1, num - 1);
			_ws = address.Substring(num + 1);
		}
		else
		{
			_wb = "";
			_ws = address;
		}
		if (_ws.StartsWith("'"))
		{
			num = _ws.IndexOf("'", 1);
			while (num > 0 && num + 1 < _ws.Length && _ws[num + 1] == '\'')
			{
				_ws = _ws.Substring(0, num) + _ws.Substring(num + 1);
				num = _ws.IndexOf("'", num + 1);
			}
			if (num > 0)
			{
				_address = _ws.Substring(num + 2);
				_ws = _ws.Substring(1, num - 1);
				return;
			}
		}
		num = _ws.IndexOf("!");
		if (num == 0)
		{
			_address = _ws.Substring(1);
			_ws = _wb;
			_wb = "";
		}
		else if (num > -1)
		{
			_address = _ws.Substring(num + 1);
			_ws = _ws.Substring(0, num);
		}
		else
		{
			_address = address;
		}
	}

	internal void ChangeWorksheet(string wsName, string newWs)
	{
		if (_ws == wsName)
		{
			_ws = newWs;
		}
		string text = GetAddress();
		if (Addresses != null)
		{
			foreach (ExcelAddress address in Addresses)
			{
				if (address._ws == wsName)
				{
					address._ws = newWs;
					text = text + "," + address.GetAddress();
				}
				else
				{
					text = text + "," + address._address;
				}
			}
		}
		_address = text;
	}

	private string GetAddress()
	{
		string text = "";
		if (!string.IsNullOrEmpty(_wb))
		{
			text = "[" + _wb + "]";
		}
		if (!string.IsNullOrEmpty(_ws))
		{
			text += $"'{_ws}'!";
		}
		if (IsName)
		{
			return text + ExcelCellBase.GetAddress(_fromRow, _fromCol, _toRow, _toCol);
		}
		return text + ExcelCellBase.GetAddress(_fromRow, _fromCol, _toRow, _toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
	}

	public override string ToString()
	{
		return _address;
	}

	protected void Validate()
	{
		if (_fromRow > _toRow || _fromCol > _toCol)
		{
			throw new ArgumentOutOfRangeException("Start cell Address must be less or equal to End cell address");
		}
	}

	private bool ExtractAddress(string fullAddress)
	{
		Stack<int> stack = new Stack<int>();
		List<string> list = new List<string>();
		string first = "";
		string second = "";
		bool flag = false;
		bool hasSheet = false;
		try
		{
			if (fullAddress == "#REF!")
			{
				SetAddress(ref fullAddress, ref second, ref hasSheet, isMulti: false);
				return true;
			}
			if (ConvertUtil._invariantCompareInfo.IsPrefix(fullAddress, "!"))
			{
				return false;
			}
			bool isMulti = false;
			for (int i = 0; i < fullAddress.Length; i++)
			{
				char c = fullAddress[i];
				if (c == '\'')
				{
					if (flag && i + 1 < fullAddress.Length && fullAddress[i] == '\'')
					{
						if (hasSheet)
						{
							second += c;
						}
						else
						{
							first += c;
						}
					}
					flag = !flag;
				}
				else if (stack.Count > 0)
				{
					if (c == '[' && !flag)
					{
						stack.Push(i);
					}
					else if (c == ']' && !flag)
					{
						if (stack.Count <= 0)
						{
							return false;
						}
						int num = stack.Pop();
						list.Add(fullAddress.Substring(num + 1, i - num - 1));
						if (stack.Count == 0)
						{
							HandleBrackets(first, second, list);
						}
					}
				}
				else if (c == '[' && !flag)
				{
					stack.Push(i);
				}
				else if (c == '!' && !flag && !first.EndsWith("#REF") && !second.EndsWith("#REF"))
				{
					if (hasSheet && second != null && second.ToLower().EndsWith(first.ToLower()))
					{
						second = Regex.Replace(second, $"{first}$", string.Empty);
					}
					hasSheet = true;
				}
				else if (c == ',' && !flag)
				{
					isMulti = true;
					SetAddress(ref first, ref second, ref hasSheet, isMulti);
				}
				else if (hasSheet)
				{
					second += c;
				}
				else
				{
					first += c;
				}
			}
			if (Table == null)
			{
				SetAddress(ref first, ref second, ref hasSheet, isMulti);
			}
			return true;
		}
		catch
		{
			return false;
		}
	}

	private void HandleBrackets(string first, string second, List<string> bracketParts)
	{
		if (string.IsNullOrEmpty(first))
		{
			return;
		}
		_table = new ExcelTableAddress();
		Table.Name = first;
		foreach (string bracketPart in bracketParts)
		{
			if (bracketPart.IndexOf("[") >= 0)
			{
				continue;
			}
			switch (bracketPart.ToLower(CultureInfo.InvariantCulture))
			{
			case "#all":
				_table.IsAll = true;
				continue;
			case "#headers":
				_table.IsHeader = true;
				continue;
			case "#data":
				_table.IsData = true;
				continue;
			case "#totals":
				_table.IsTotals = true;
				continue;
			case "#this row":
				_table.IsThisRow = true;
				continue;
			}
			if (string.IsNullOrEmpty(_table.ColumnSpan))
			{
				_table.ColumnSpan = bracketPart;
				continue;
			}
			ExcelTableAddress table = _table;
			table.ColumnSpan = table.ColumnSpan + ":" + bracketPart;
		}
	}

	internal eAddressCollition Collide(ExcelAddressBase address, bool ignoreWs = false)
	{
		if (!ignoreWs && address.WorkSheet != WorkSheet && address.WorkSheet != null)
		{
			return eAddressCollition.No;
		}
		if (address._fromRow > _toRow || address._fromCol > _toCol || _fromRow > address._toRow || _fromCol > address._toCol)
		{
			return eAddressCollition.No;
		}
		if (address._fromRow == _fromRow && address._fromCol == _fromCol && address._toRow == _toRow && address._toCol == _toCol)
		{
			return eAddressCollition.Equal;
		}
		if (address._fromRow >= _fromRow && address._toRow <= _toRow && address._fromCol >= _fromCol && address._toCol <= _toCol)
		{
			return eAddressCollition.Inside;
		}
		return eAddressCollition.Partly;
	}

	internal ExcelAddressBase AddRow(int row, int rows, bool setFixed = false)
	{
		if (row > _toRow)
		{
			return this;
		}
		if (row <= _fromRow)
		{
			return new ExcelAddressBase((setFixed && _fromRowFixed) ? _fromRow : (_fromRow + rows), _fromCol, (setFixed && _toRowFixed) ? _toRow : (_toRow + rows), _toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
		}
		return new ExcelAddressBase(_fromRow, _fromCol, (setFixed && _toRowFixed) ? _toRow : (_toRow + rows), _toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
	}

	internal ExcelAddressBase DeleteRow(int row, int rows, bool setFixed = false)
	{
		if (row > _toRow)
		{
			return this;
		}
		if (row + rows <= _fromRow)
		{
			return new ExcelAddressBase((setFixed && _fromRowFixed) ? _fromRow : (_fromRow - rows), _fromCol, (setFixed && _toRowFixed) ? _toRow : (_toRow - rows), _toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
		}
		if (row <= _fromRow && row + rows > _toRow)
		{
			return null;
		}
		if (row <= _fromRow)
		{
			return new ExcelAddressBase(row, _fromCol, (setFixed && _toRowFixed) ? _toRow : (_toRow - rows), _toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
		}
		return new ExcelAddressBase(_fromRow, _fromCol, (setFixed && _toRowFixed) ? _toRow : ((_toRow - rows < row) ? (row - 1) : (_toRow - rows)), _toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
	}

	internal ExcelAddressBase AddColumn(int col, int cols, bool setFixed = false)
	{
		if (col > _toCol)
		{
			return this;
		}
		if (col <= _fromCol)
		{
			return new ExcelAddressBase(_fromRow, (setFixed && _fromColFixed) ? _fromCol : (_fromCol + cols), _toRow, (setFixed && _toColFixed) ? _toCol : (_toCol + cols), _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
		}
		return new ExcelAddressBase(_fromRow, _fromCol, _toRow, (setFixed && _toColFixed) ? _toCol : (_toCol + cols), _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
	}

	internal ExcelAddressBase DeleteColumn(int col, int cols, bool setFixed = false)
	{
		if (col > _toCol)
		{
			return this;
		}
		if (col + cols <= _fromCol)
		{
			return new ExcelAddressBase(_fromRow, (setFixed && _fromColFixed) ? _fromCol : (_fromCol - cols), _toRow, (setFixed && _toColFixed) ? _toCol : (_toCol - cols), _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
		}
		if (col <= _fromCol && col + cols > _toCol)
		{
			return null;
		}
		if (col <= _fromCol)
		{
			return new ExcelAddressBase(_fromRow, col, _toRow, (setFixed && _toColFixed) ? _toCol : (_toCol - cols), _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
		}
		return new ExcelAddressBase(_fromRow, _fromCol, _toRow, (setFixed && _toColFixed) ? _toCol : ((_toCol - cols < col) ? (col - 1) : (_toCol - cols)), _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
	}

	internal ExcelAddressBase Insert(ExcelAddressBase address, eShiftType Shift)
	{
		if (_toRow < address._fromRow || _toCol < address._fromCol || (_fromRow > address._toRow && _fromCol > address._toCol))
		{
			return this;
		}
		_ = address.Rows;
		_ = address.Columns;
		if (Shift == eShiftType.Right)
		{
			if (address._fromRow > _fromRow)
			{
				ExcelCellBase.GetAddress(_fromRow, _fromCol, address._fromRow, _toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
			}
			if (address._fromCol > _fromCol)
			{
				ExcelCellBase.GetAddress((_fromRow < address._fromRow) ? _fromRow : address._fromRow, _fromCol, address._fromRow, _toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
			}
		}
		if (_toRow < address._fromRow)
		{
			_ = _fromRow;
			_ = address._fromRow;
		}
		return null;
	}

	private void SetAddress(ref string first, ref string second, ref bool hasSheet, bool isMulti)
	{
		string text;
		string text2;
		if (hasSheet)
		{
			text = first;
			text2 = second;
			first = "";
			second = "";
		}
		else
		{
			text2 = first;
			text = "";
			first = "";
		}
		hasSheet = false;
		if (string.IsNullOrEmpty(_firstAddress))
		{
			if (string.IsNullOrEmpty(_ws) || !string.IsNullOrEmpty(text))
			{
				_ws = text;
			}
			_firstAddress = text2;
			ExcelCellBase.GetRowColFromAddress(text2, out _fromRow, out _fromCol, out _toRow, out _toCol, out _fromRowFixed, out _fromColFixed, out _toRowFixed, out _toColFixed);
		}
		if (isMulti)
		{
			if (_addresses == null)
			{
				_addresses = new List<ExcelAddress>();
			}
			_addresses.Add(new ExcelAddress(_ws, text2));
		}
		else
		{
			_addresses = null;
		}
	}

	internal static AddressType IsValid(string Address, bool r1c1 = false)
	{
		if (Address == "#REF!")
		{
			return AddressType.Invalid;
		}
		if (double.TryParse(Address, NumberStyles.Any, CultureInfo.InvariantCulture, out var _))
		{
			return AddressType.Invalid;
		}
		if (IsFormula(Address))
		{
			return AddressType.Formula;
		}
		if (r1c1 && IsR1C1(Address))
		{
			return AddressType.R1C1;
		}
		if (SplitAddress(Address, out var wb, out var _, out var intAddress))
		{
			if (intAddress.Contains("["))
			{
				if (!string.IsNullOrEmpty(wb))
				{
					return AddressType.ExternalAddress;
				}
				return AddressType.InternalAddress;
			}
			if (intAddress.Contains(","))
			{
				intAddress = intAddress.Substring(0, intAddress.IndexOf(','));
			}
			if (IsAddress(intAddress))
			{
				if (!string.IsNullOrEmpty(wb))
				{
					return AddressType.ExternalAddress;
				}
				return AddressType.InternalAddress;
			}
			if (!string.IsNullOrEmpty(wb))
			{
				return AddressType.ExternalName;
			}
			return AddressType.InternalName;
		}
		return AddressType.Invalid;
	}

	private static bool IsR1C1(string address)
	{
		if (address.StartsWith("!"))
		{
			address = address.Substring(1);
		}
		address = address.ToUpper();
		if (string.IsNullOrEmpty(address) || (address[0] != 'R' && address[0] != 'C'))
		{
			return false;
		}
		bool flag = false;
		bool flag2 = false;
		bool flag3 = false;
		string text = address;
		foreach (char c in text)
		{
			switch (c)
			{
			case 'C':
				flag = true;
				flag2 = true;
				continue;
			case 'R':
				if (flag)
				{
					return false;
				}
				flag2 = true;
				continue;
			case '[':
				flag3 = true;
				continue;
			case ']':
				if (!flag3)
				{
					return false;
				}
				flag2 = false;
				continue;
			case ':':
				flag = false;
				flag3 = false;
				flag2 = false;
				continue;
			}
			if ((c >= '0' && c <= '9') || c == '-')
			{
				if (!flag2)
				{
					return false;
				}
				continue;
			}
			return false;
		}
		return true;
	}

	private static bool IsAddress(string intAddress)
	{
		if (string.IsNullOrEmpty(intAddress))
		{
			return false;
		}
		string[] array = intAddress.Split(':');
		if (!ExcelCellBase.GetRowCol(array[0], out var row, out var col, throwException: false))
		{
			return false;
		}
		int row2;
		int col2;
		if (array.Length > 1)
		{
			if (!ExcelCellBase.GetRowCol(array[1], out row2, out col2, throwException: false))
			{
				return false;
			}
		}
		else
		{
			row2 = row;
			col2 = col;
		}
		if (row <= row2 && col <= col2 && col > -1 && col2 <= 16384 && row > -1 && row2 <= 1048576)
		{
			return true;
		}
		return false;
	}

	private static bool SplitAddress(string Address, out string wb, out string ws, out string intAddress)
	{
		wb = "";
		ws = "";
		intAddress = "";
		string text = "";
		bool flag = false;
		int num = -1;
		for (int i = 0; i < Address.Length; i++)
		{
			if (Address[i] == '\'')
			{
				flag = !flag;
				if (i > 0 && Address[i - 1] == '\'')
				{
					text += "'";
				}
				continue;
			}
			if (Address[i] == '!' && !flag)
			{
				if (text.Length > 0 && text[0] == '[')
				{
					wb = text.Substring(1, text.IndexOf("]") - 1);
					ws = text.Substring(text.IndexOf("]") + 1);
				}
				else
				{
					ws = text;
				}
				intAddress = Address.Substring(i + 1);
				return true;
			}
			if (Address[i] == '[' && !flag)
			{
				if (i > 0)
				{
					intAddress = Address;
					return true;
				}
				num = i;
			}
			else if (Address[i] == ']' && !flag)
			{
				if (num <= -1)
				{
					return false;
				}
				wb = text;
				text = "";
			}
			else
			{
				text += Address[i];
			}
		}
		intAddress = text;
		return true;
	}

	private static bool IsFormula(string address)
	{
		bool flag = false;
		foreach (char c in address)
		{
			switch (c)
			{
			case '\'':
				flag = !flag;
				continue;
			case '[':
			case ']':
				return false;
			}
			if (!flag && FormulaCharacters.Contains(c))
			{
				return true;
			}
		}
		return false;
	}

	private static bool IsValidName(string address)
	{
		if (Regex.IsMatch(address, "[^0-9./*-+,½!\"@#£%&/{}()\\[\\]=?`^~':;<>|][^/*-+,½!\"@#£%&/{}()\\[\\]=?`^~':;<>|]*"))
		{
			return true;
		}
		return false;
	}

	internal static string GetWorkbookPart(string address)
	{
		int num = 0;
		if (address[0] == '[')
		{
			num = address.IndexOf(']') + 1;
			if (num > 0)
			{
				return address.Substring(1, num - 2);
			}
		}
		return "";
	}

	internal static string GetWorksheetPart(string address, string defaultWorkSheet)
	{
		int endIx = 0;
		return GetWorksheetPart(address, defaultWorkSheet, ref endIx);
	}

	internal static string GetWorksheetPart(string address, string defaultWorkSheet, ref int endIx)
	{
		if (address == "")
		{
			return defaultWorkSheet;
		}
		int num = 0;
		if (address[0] == '[')
		{
			num = address.IndexOf(']') + 1;
		}
		if (num > 0 && num < address.Length)
		{
			if (address[num] == '\'')
			{
				return GetString(address, num, out endIx);
			}
			int num2 = address.IndexOf('!', num);
			if (num2 > num)
			{
				return address.Substring(num, num2 - num);
			}
			return defaultWorkSheet;
		}
		return defaultWorkSheet;
	}

	internal static string GetAddressPart(string address)
	{
		int endIx = 0;
		GetWorksheetPart(address, "", ref endIx);
		if (endIx < address.Length)
		{
			if (address[endIx] == '!')
			{
				return address.Substring(endIx + 1);
			}
			return "";
		}
		return "";
	}

	internal static void SplitAddress(string fullAddress, out string wb, out string ws, out string address, string defaultWorksheet = "")
	{
		wb = GetWorkbookPart(fullAddress);
		int endIx = 0;
		ws = GetWorksheetPart(fullAddress, defaultWorksheet, ref endIx);
		if (endIx < fullAddress.Length)
		{
			if (fullAddress[endIx] == '!')
			{
				address = fullAddress.Substring(endIx + 1);
			}
			else
			{
				address = fullAddress.Substring(endIx);
			}
		}
		else
		{
			address = "";
		}
	}

	private static string GetString(string address, int ix, out int endIx)
	{
		for (int num = address.IndexOf("''"); num > -1; num = address.IndexOf("''"))
		{
		}
		endIx = address.IndexOf("'");
		return address.Substring(ix, endIx - ix).Replace("''", "'");
	}

	internal bool IsValidRowCol()
	{
		if (_fromRow <= _toRow && _fromCol <= _toCol && _fromRow >= 1 && _fromCol >= 1 && _toRow <= 1048576)
		{
			return _toCol <= 16384;
		}
		return false;
	}
}
