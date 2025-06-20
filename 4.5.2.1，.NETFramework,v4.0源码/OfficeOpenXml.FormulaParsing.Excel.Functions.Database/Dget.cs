using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Database;

public class Dget : DatabaseFunction
{
	public Dget()
		: this(new RowMatcher())
	{
	}

	public Dget(RowMatcher rowMatcher)
		: base(rowMatcher)
	{
	}

	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 3);
		string address = arguments.ElementAt(0).ValueAsRangeInfo.Address.Address;
		string field = ArgToString(arguments, 1).ToLower(CultureInfo.InvariantCulture);
		string address2 = arguments.ElementAt(2).ValueAsRangeInfo.Address.Address;
		ExcelDatabase excelDatabase = new ExcelDatabase(context.ExcelDataProvider, address);
		ExcelDatabaseCriteria criteria = new ExcelDatabaseCriteria(context.ExcelDataProvider, address2);
		int num = 0;
		object obj = null;
		while (excelDatabase.HasMoreRows)
		{
			ExcelDatabaseRow excelDatabaseRow = excelDatabase.Read();
			if (base.RowMatcher.IsMatch(excelDatabaseRow, criteria))
			{
				if (++num > 1)
				{
					return CreateResult("#NUM!", DataType.ExcelError);
				}
				obj = excelDatabaseRow[field];
			}
		}
		return new CompileResultFactory().Create(obj);
	}
}
