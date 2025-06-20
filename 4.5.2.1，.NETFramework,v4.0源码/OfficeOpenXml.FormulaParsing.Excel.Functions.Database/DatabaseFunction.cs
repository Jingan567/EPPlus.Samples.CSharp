using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Database;

public abstract class DatabaseFunction : ExcelFunction
{
	protected RowMatcher RowMatcher { get; private set; }

	public DatabaseFunction()
		: this(new RowMatcher())
	{
	}

	public DatabaseFunction(RowMatcher rowMatcher)
	{
		RowMatcher = rowMatcher;
	}

	protected IEnumerable<double> GetMatchingValues(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		string address = arguments.ElementAt(0).ValueAsRangeInfo.Address.Address;
		object value = arguments.ElementAt(1).Value;
		string address2 = arguments.ElementAt(2).ValueAsRangeInfo.Address.Address;
		ExcelDatabase excelDatabase = new ExcelDatabase(context.ExcelDataProvider, address);
		ExcelDatabaseCriteria criteria = new ExcelDatabaseCriteria(context.ExcelDataProvider, address2);
		List<double> list = new List<double>();
		while (excelDatabase.HasMoreRows)
		{
			ExcelDatabaseRow excelDatabaseRow = excelDatabase.Read();
			if (RowMatcher.IsMatch(excelDatabaseRow, criteria))
			{
				object obj = (ConvertUtil.IsNumeric(value) ? excelDatabaseRow[(int)ConvertUtil.GetValueDouble(value)] : excelDatabaseRow[value.ToString().ToLower(CultureInfo.InvariantCulture)]);
				if (ConvertUtil.IsNumeric(obj))
				{
					list.Add(ConvertUtil.GetValueDouble(obj));
				}
			}
		}
		return list;
	}
}
