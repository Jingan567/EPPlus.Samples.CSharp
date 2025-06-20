using System;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Information;

public class IsOdd : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		object firstValue = GetFirstValue(arguments);
		if (!ConvertUtil.IsNumeric(firstValue))
		{
			ThrowExcelErrorValueException(eErrorType.Value);
		}
		int num = (int)System.Math.Floor(ConvertUtil.GetValueDouble(firstValue));
		return CreateResult(num % 2 == 1, DataType.Boolean);
	}
}
