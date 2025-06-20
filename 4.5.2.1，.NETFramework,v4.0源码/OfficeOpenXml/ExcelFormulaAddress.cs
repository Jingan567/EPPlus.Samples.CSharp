using System.Collections.Generic;

namespace OfficeOpenXml;

public class ExcelFormulaAddress : ExcelAddressBase
{
	internal new List<ExcelFormulaAddress> _addresses;

	public new string Address
	{
		get
		{
			if (string.IsNullOrEmpty(_address) && _fromRow > 0)
			{
				_address = ExcelCellBase.GetAddress(_fromRow, _fromCol, _toRow, _toCol, _fromRowFixed, _toRowFixed, _fromColFixed, _toColFixed);
			}
			return _address;
		}
		set
		{
			SetAddress(value);
			ChangeAddress();
			SetFixed();
		}
	}

	public new List<ExcelFormulaAddress> Addresses
	{
		get
		{
			if (_addresses == null)
			{
				_addresses = new List<ExcelFormulaAddress>();
			}
			return _addresses;
		}
	}

	internal ExcelFormulaAddress()
	{
	}

	public ExcelFormulaAddress(int fromRow, int fromCol, int toRow, int toColumn)
		: base(fromRow, fromCol, toRow, toColumn)
	{
		_ws = "";
	}

	public ExcelFormulaAddress(string address)
		: base(address)
	{
		SetFixed();
	}

	internal ExcelFormulaAddress(string ws, string address)
		: base(address)
	{
		if (string.IsNullOrEmpty(_ws))
		{
			_ws = ws;
		}
		SetFixed();
	}

	internal ExcelFormulaAddress(string ws, string address, bool isName)
		: base(address, isName)
	{
		if (string.IsNullOrEmpty(_ws))
		{
			_ws = ws;
		}
		if (!isName)
		{
			SetFixed();
		}
	}

	private void SetFixed()
	{
		if (Address.IndexOf("[") < 0)
		{
			string firstAddress = base.FirstAddress;
			if (_fromRow == _toRow && _fromCol == _toCol)
			{
				GetFixed(firstAddress, out _fromRowFixed, out _fromColFixed);
				return;
			}
			string[] array = firstAddress.Split(':');
			GetFixed(array[0], out _fromRowFixed, out _fromColFixed);
			GetFixed(array[1], out _toRowFixed, out _toColFixed);
		}
	}

	private void GetFixed(string address, out bool rowFixed, out bool colFixed)
	{
		rowFixed = (colFixed = false);
		int num;
		for (num = address.IndexOf('$'); num > -1; num = address.IndexOf('$', num))
		{
			num++;
			if (num < address.Length)
			{
				if (address[num] >= '0' && address[num] <= '9')
				{
					rowFixed = true;
					break;
				}
				colFixed = true;
			}
		}
	}

	internal string GetOffset(int row, int column)
	{
		int num = _fromRow;
		int num2 = _fromCol;
		int num3 = _toRow;
		int num4 = _toCol;
		bool num5 = num != num3 || num2 != num4;
		if (!_fromRowFixed)
		{
			num += row;
		}
		if (!_fromColFixed)
		{
			num2 += column;
		}
		if (num5)
		{
			if (!_toRowFixed)
			{
				num3 += row;
			}
			if (!_toColFixed)
			{
				num4 += column;
			}
		}
		else
		{
			num3 = num;
			num4 = num2;
		}
		string text = ExcelCellBase.GetAddress(num, num2, num3, num4, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
		if (Addresses != null)
		{
			foreach (ExcelFormulaAddress address in Addresses)
			{
				text = text + "," + address.GetOffset(row, column);
			}
		}
		return text;
	}
}
