using System;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;

public class Date : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 3);
		int year = ArgToInt(arguments, 0);
		int num = ArgToInt(arguments, 1);
		int num2 = ArgToInt(arguments, 2);
		System.DateTime dateTime = new System.DateTime(year, 1, 1);
		num--;
		return CreateResult(dateTime.AddMonths(num).AddDays(num2 - 1).ToOADate(), DataType.Date);
	}
}
