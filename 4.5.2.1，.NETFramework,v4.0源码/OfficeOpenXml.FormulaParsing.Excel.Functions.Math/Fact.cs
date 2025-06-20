using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class Fact : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		double number = ArgToDecimal(arguments, 0);
		ThrowExcelErrorValueExceptionIf(() => number < 0.0, eErrorType.NA);
		double num = 1.0;
		for (int i = 1; (double)i < number; i++)
		{
			num *= (double)i;
		}
		return CreateResult(num, DataType.Integer);
	}
}
