using System.Collections.Generic;
using System.Globalization;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Database;

public class ExcelDatabaseCriteria
{
	private readonly ExcelDataProvider _dataProvider;

	private readonly int _fromCol;

	private readonly int _toCol;

	private readonly string _worksheet;

	private readonly int _fieldRow;

	private readonly Dictionary<ExcelDatabaseCriteriaField, object> _criterias = new Dictionary<ExcelDatabaseCriteriaField, object>();

	public virtual IDictionary<ExcelDatabaseCriteriaField, object> Items => _criterias;

	public ExcelDatabaseCriteria(ExcelDataProvider dataProvider, string range)
	{
		_dataProvider = dataProvider;
		ExcelAddressBase excelAddressBase = new ExcelAddressBase(range);
		_fromCol = excelAddressBase._fromCol;
		_toCol = excelAddressBase._toCol;
		_worksheet = excelAddressBase.WorkSheet;
		_fieldRow = excelAddressBase._fromRow;
		Initialize();
	}

	private void Initialize()
	{
		for (int i = _fromCol; i <= _toCol; i++)
		{
			object cellValue = _dataProvider.GetCellValue(_worksheet, _fieldRow, i);
			object cellValue2 = _dataProvider.GetCellValue(_worksheet, _fieldRow + 1, i);
			if (cellValue != null && cellValue2 != null)
			{
				if (cellValue is string)
				{
					ExcelDatabaseCriteriaField key = new ExcelDatabaseCriteriaField(cellValue.ToString().ToLower(CultureInfo.InvariantCulture));
					_criterias.Add(key, cellValue2);
				}
				else if (ConvertUtil.IsNumeric(cellValue))
				{
					ExcelDatabaseCriteriaField key2 = new ExcelDatabaseCriteriaField((int)cellValue);
					_criterias.Add(key2, cellValue2);
				}
			}
		}
	}
}
