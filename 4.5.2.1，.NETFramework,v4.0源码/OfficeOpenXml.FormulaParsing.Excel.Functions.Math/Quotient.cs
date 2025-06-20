using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class Quotient : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 2);
		double num = ArgToDecimal(arguments, 0);
		double denom = ArgToDecimal(arguments, 1);
		ThrowExcelErrorValueExceptionIf(() => (int)denom == 0, eErrorType.Div0);
		int num2 = (int)(num / denom);
		return CreateResult(num2, DataType.Integer);
	}
}
