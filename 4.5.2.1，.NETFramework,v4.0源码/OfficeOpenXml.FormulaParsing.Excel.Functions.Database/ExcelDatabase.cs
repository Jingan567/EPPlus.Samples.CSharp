using System.Collections.Generic;
using System.Globalization;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Database;

public class ExcelDatabase
{
	private readonly ExcelDataProvider _dataProvider;

	private readonly int _fromCol;

	private readonly int _toCol;

	private readonly int _fieldRow;

	private readonly int _endRow;

	private readonly string _worksheet;

	private int _rowIndex;

	private readonly List<ExcelDatabaseField> _fields = new List<ExcelDatabaseField>();

	public IEnumerable<ExcelDatabaseField> Fields => _fields;

	public bool HasMoreRows => _rowIndex < _endRow;

	public ExcelDatabase(ExcelDataProvider dataProvider, string range)
	{
		_dataProvider = dataProvider;
		ExcelAddressBase excelAddressBase = new ExcelAddressBase(range);
		_fromCol = excelAddressBase._fromCol;
		_toCol = excelAddressBase._toCol;
		_fieldRow = excelAddressBase._fromRow;
		_endRow = excelAddressBase._toRow;
		_worksheet = excelAddressBase.WorkSheet;
		_rowIndex = _fieldRow;
		Initialize();
	}

	private void Initialize()
	{
		int num = 0;
		for (int i = _fromCol; i <= _toCol; i++)
		{
			object cellValue = GetCellValue(_fieldRow, i);
			string fieldName = ((cellValue != null) ? cellValue.ToString().ToLower(CultureInfo.InvariantCulture) : string.Empty);
			_fields.Add(new ExcelDatabaseField(fieldName, num++));
		}
	}

	private object GetCellValue(int row, int col)
	{
		return _dataProvider.GetRangeValue(_worksheet, row, col);
	}

	public ExcelDatabaseRow Read()
	{
		ExcelDatabaseRow excelDatabaseRow = new ExcelDatabaseRow();
		_rowIndex++;
		foreach (ExcelDatabaseField field in Fields)
		{
			int col = _fromCol + field.ColIndex;
			object cellValue = GetCellValue(_rowIndex, col);
			excelDatabaseRow[field.FieldName] = cellValue;
		}
		return excelDatabaseRow;
	}
}
