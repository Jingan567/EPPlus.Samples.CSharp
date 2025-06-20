using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities;

public class RangeAddressFactory
{
	private readonly ExcelDataProvider _excelDataProvider;

	private readonly AddressTranslator _addressTranslator;

	private readonly IndexToAddressTranslator _indexToAddressTranslator;

	public RangeAddressFactory(ExcelDataProvider excelDataProvider)
		: this(excelDataProvider, new AddressTranslator(excelDataProvider), new IndexToAddressTranslator(excelDataProvider, ExcelReferenceType.RelativeRowAndColumn))
	{
	}

	public RangeAddressFactory(ExcelDataProvider excelDataProvider, AddressTranslator addressTranslator, IndexToAddressTranslator indexToAddressTranslator)
	{
		Require.That(excelDataProvider).Named("excelDataProvider").IsNotNull();
		Require.That(addressTranslator).Named("addressTranslator").IsNotNull();
		Require.That(indexToAddressTranslator).Named("indexToAddressTranslator").IsNotNull();
		_excelDataProvider = excelDataProvider;
		_addressTranslator = addressTranslator;
		_indexToAddressTranslator = indexToAddressTranslator;
	}

	public RangeAddress Create(int col, int row)
	{
		return Create(string.Empty, col, row);
	}

	public RangeAddress Create(string worksheetName, int col, int row)
	{
		return new RangeAddress
		{
			Address = _indexToAddressTranslator.ToAddress(col, row),
			Worksheet = worksheetName,
			FromCol = col,
			ToCol = col,
			FromRow = row,
			ToRow = row
		};
	}

	public RangeAddress Create(string worksheetName, string address)
	{
		Require.That(address).Named("range").IsNotNullOrEmpty();
		ExcelAddressBase excelAddressBase = new ExcelAddressBase(address);
		string worksheet = (string.IsNullOrEmpty(excelAddressBase.WorkSheet) ? worksheetName : excelAddressBase.WorkSheet);
		ExcelCellAddress dimensionEnd = _excelDataProvider.GetDimensionEnd(worksheet);
		return new RangeAddress
		{
			Address = excelAddressBase.Address,
			Worksheet = worksheet,
			FromRow = excelAddressBase._fromRow,
			FromCol = excelAddressBase._fromCol,
			ToRow = ((dimensionEnd != null && excelAddressBase._toRow > dimensionEnd.Row) ? dimensionEnd.Row : excelAddressBase._toRow),
			ToCol = excelAddressBase._toCol
		};
	}

	public RangeAddress Create(string range)
	{
		Require.That(range).Named("range").IsNotNullOrEmpty();
		ExcelAddressBase excelAddressBase = new ExcelAddressBase(range);
		if (excelAddressBase.Table != null)
		{
			ExcelAddressBase address = _excelDataProvider.GetRange(excelAddressBase.WorkSheet, range).Address;
			excelAddressBase = new ExcelAddressBase(address._fromRow, address._fromCol, address._toRow, address._toCol);
			excelAddressBase._ws = address._ws;
		}
		return new RangeAddress
		{
			Address = excelAddressBase.Address,
			Worksheet = (excelAddressBase.WorkSheet ?? ""),
			FromRow = excelAddressBase._fromRow,
			FromCol = excelAddressBase._fromCol,
			ToRow = excelAddressBase._toRow,
			ToCol = excelAddressBase._toCol
		};
	}

	private void HandleSingleCellAddress(RangeAddress rangeAddress, ExcelAddressInfo addressInfo)
	{
		_addressTranslator.ToColAndRow(addressInfo.StartCell, out var col, out var row);
		rangeAddress.FromCol = col;
		rangeAddress.ToCol = col;
		rangeAddress.FromRow = row;
		rangeAddress.ToRow = row;
	}

	private void HandleMultipleCellAddress(RangeAddress rangeAddress, ExcelAddressInfo addressInfo)
	{
		_addressTranslator.ToColAndRow(addressInfo.StartCell, out var col, out var row);
		_addressTranslator.ToColAndRow(addressInfo.EndCell, out var col2, out var row2, AddressTranslator.RangeCalculationBehaviour.LastPart);
		rangeAddress.FromCol = col;
		rangeAddress.ToCol = col2;
		rangeAddress.FromRow = row;
		rangeAddress.ToRow = row2;
	}
}
