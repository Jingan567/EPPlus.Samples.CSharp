using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Database;

public class Dcount : ExcelFunction
{
	private readonly RowMatcher _rowMatcher;

	public Dcount()
		: this(new RowMatcher())
	{
	}

	public Dcount(RowMatcher rowMatcher)
	{
		_rowMatcher = rowMatcher;
	}

	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 2);
		string address = arguments.ElementAt(0).ValueAsRangeInfo.Address.Address;
		string text = null;
		string text2 = null;
		if (arguments.Count() == 2)
		{
			text2 = arguments.ElementAt(1).ValueAsRangeInfo.Address.Address;
		}
		else
		{
			text = ArgToString(arguments, 1).ToLower(CultureInfo.InvariantCulture);
			text2 = arguments.ElementAt(2).ValueAsRangeInfo.Address.Address;
		}
		ExcelDatabase excelDatabase = new ExcelDatabase(context.ExcelDataProvider, address);
		ExcelDatabaseCriteria criteria = new ExcelDatabaseCriteria(context.ExcelDataProvider, text2);
		int num = 0;
		while (excelDatabase.HasMoreRows)
		{
			ExcelDatabaseRow excelDatabaseRow = excelDatabase.Read();
			if (!_rowMatcher.IsMatch(excelDatabaseRow, criteria))
			{
				continue;
			}
			if (!string.IsNullOrEmpty(text))
			{
				if (ConvertUtil.IsNumeric(excelDatabaseRow[text]))
				{
					num++;
				}
			}
			else
			{
				num++;
			}
		}
		return CreateResult(num, DataType.Integer);
	}
}
