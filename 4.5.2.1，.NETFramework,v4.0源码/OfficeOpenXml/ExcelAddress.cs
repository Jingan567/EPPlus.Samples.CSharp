namespace OfficeOpenXml;

public class ExcelAddress : ExcelAddressBase
{
	public new string Address
	{
		get
		{
			if (string.IsNullOrEmpty(_address) && _fromRow > 0)
			{
				_address = ExcelCellBase.GetAddress(_fromRow, _fromCol, _toRow, _toCol);
			}
			return _address;
		}
		set
		{
			SetAddress(value);
			ChangeAddress();
		}
	}

	internal ExcelAddress()
	{
	}

	public ExcelAddress(int fromRow, int fromCol, int toRow, int toColumn)
		: base(fromRow, fromCol, toRow, toColumn)
	{
		_ws = "";
	}

	public ExcelAddress(string address)
		: base(address)
	{
	}

	internal ExcelAddress(string ws, string address)
		: base(address)
	{
		if (string.IsNullOrEmpty(_ws))
		{
			_ws = ws;
		}
	}

	internal ExcelAddress(string ws, string address, bool isName)
		: base(address, isName)
	{
		if (string.IsNullOrEmpty(_ws))
		{
			_ws = ws;
		}
	}

	public ExcelAddress(string Address, ExcelPackage package, ExcelAddressBase referenceAddress)
		: base(Address, package, referenceAddress)
	{
	}
}
