using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

public class Search : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		FunctionArgument[] array = (arguments as FunctionArgument[]) ?? arguments.ToArray();
		ValidateArguments(array, 2);
		string value = ArgToString(array, 0);
		string text = ArgToString(array, 1);
		int startIndex = 0;
		if (array.Count() > 2)
		{
			startIndex = ArgToInt(array, 2);
		}
		int num = text.IndexOf(value, startIndex, StringComparison.OrdinalIgnoreCase);
		if (num == -1)
		{
			return CreateResult(ExcelErrorValue.Create(eErrorType.Value), DataType.ExcelError);
		}
		return CreateResult(num + 1, DataType.Integer);
	}
}
