using System;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;

public class DateValue : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		string dateString = ArgToString(arguments, 0);
		return Execute(dateString);
	}

	internal CompileResult Execute(string dateString)
	{
		System.DateTime.TryParse(dateString, out var result);
		if (!(result != System.DateTime.MinValue))
		{
			return CreateResult(ExcelErrorValue.Create(eErrorType.Value), DataType.ExcelError);
		}
		return CreateResult(result.ToOADate(), DataType.Date);
	}
}
