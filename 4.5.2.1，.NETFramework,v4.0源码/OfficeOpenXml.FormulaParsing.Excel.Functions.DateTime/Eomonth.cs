using System;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;

public class Eomonth : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 2);
		System.DateTime dateTime = System.DateTime.FromOADate(ArgToDecimal(arguments, 0));
		int num = ArgToInt(arguments, 1);
		return CreateResult(new System.DateTime(dateTime.Year, dateTime.Month, 1).AddMonths(num + 1).AddDays(-1.0).ToOADate(), DataType.Date);
	}
}
