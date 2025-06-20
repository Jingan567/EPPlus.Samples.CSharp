using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Database;

public class DcountA : DatabaseFunction
{
	public DcountA()
		: this(new RowMatcher())
	{
	}

	public DcountA(RowMatcher rowMatcher)
		: base(rowMatcher)
	{
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
			if (!base.RowMatcher.IsMatch(excelDatabaseRow, criteria))
			{
				continue;
			}
			if (!string.IsNullOrEmpty(text))
			{
				object value = excelDatabaseRow[text];
				if (ShouldCount(value))
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

	private bool ShouldCount(object value)
	{
		if (value == null)
		{
			return false;
		}
		return !string.IsNullOrEmpty(value.ToString());
	}
}
