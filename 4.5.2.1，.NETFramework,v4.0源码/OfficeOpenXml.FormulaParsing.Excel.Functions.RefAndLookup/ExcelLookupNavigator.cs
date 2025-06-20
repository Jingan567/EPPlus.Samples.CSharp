using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

public class ExcelLookupNavigator : LookupNavigator
{
	private int _currentRow;

	private int _currentCol;

	private object _currentValue;

	private RangeAddress _rangeAddress;

	private int _index;

	public override int Index => _index;

	public override object CurrentValue => _currentValue;

	public ExcelLookupNavigator(LookupDirection direction, LookupArguments arguments, ParsingContext parsingContext)
		: base(direction, arguments, parsingContext)
	{
		Initialize();
	}

	private void Initialize()
	{
		_index = 0;
		RangeAddressFactory rangeAddressFactory = new RangeAddressFactory(ParsingContext.ExcelDataProvider);
		if (Arguments.RangeInfo == null)
		{
			_rangeAddress = rangeAddressFactory.Create(ParsingContext.Scopes.Current.Address.Worksheet, Arguments.RangeAddress);
		}
		else
		{
			_rangeAddress = rangeAddressFactory.Create(Arguments.RangeInfo.Address.WorkSheet, Arguments.RangeInfo.Address.Address);
		}
		_currentCol = _rangeAddress.FromCol;
		_currentRow = _rangeAddress.FromRow;
		SetCurrentValue();
	}

	private void SetCurrentValue()
	{
		_currentValue = ParsingContext.ExcelDataProvider.GetCellValue(_rangeAddress.Worksheet, _currentRow, _currentCol);
	}

	private bool HasNext()
	{
		if (Direction == LookupDirection.Vertical)
		{
			return _currentRow < _rangeAddress.ToRow;
		}
		return _currentCol < _rangeAddress.ToCol;
	}

	public override bool MoveNext()
	{
		if (!HasNext())
		{
			return false;
		}
		if (Direction == LookupDirection.Vertical)
		{
			_currentRow++;
		}
		else
		{
			_currentCol++;
		}
		_index++;
		SetCurrentValue();
		return true;
	}

	public override object GetLookupValue()
	{
		int currentRow = _currentRow;
		int currentCol = _currentCol;
		if (Direction == LookupDirection.Vertical)
		{
			currentCol += Arguments.LookupIndex - 1;
			currentRow += Arguments.LookupOffset;
		}
		else
		{
			currentRow += Arguments.LookupIndex - 1;
			currentCol += Arguments.LookupOffset;
		}
		return ParsingContext.ExcelDataProvider.GetCellValue(_rangeAddress.Worksheet, currentRow, currentCol);
	}
}
