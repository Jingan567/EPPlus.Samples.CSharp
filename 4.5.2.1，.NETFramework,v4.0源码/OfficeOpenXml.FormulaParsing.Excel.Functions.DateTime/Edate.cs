using System;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;

public class Edate : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 2, eErrorType.Value);
		System.DateTime dateTime = System.DateTime.FromOADate(ArgToDecimal(arguments, 0));
		int months = ArgToInt(arguments, 1);
		return CreateResult(dateTime.AddMonths(months).ToOADate(), DataType.Date);
	}
}
