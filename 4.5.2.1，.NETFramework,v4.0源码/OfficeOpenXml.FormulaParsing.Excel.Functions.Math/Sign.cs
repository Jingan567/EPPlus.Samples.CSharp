using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class Sign : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		double num = 0.0;
		double num2 = ArgToDecimal(arguments, 0);
		if (num2 < 0.0)
		{
			num = -1.0;
		}
		else if (num2 > 0.0)
		{
			num = 1.0;
		}
		return CreateResult(num, DataType.Decimal);
	}
}
