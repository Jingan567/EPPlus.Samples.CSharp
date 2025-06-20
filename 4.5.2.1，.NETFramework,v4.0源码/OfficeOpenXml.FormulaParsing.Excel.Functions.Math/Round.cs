using System;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class Round : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 2);
		double num = ArgToDecimal(arguments, 0);
		int num2 = ArgToInt(arguments, 1);
		if (num2 < 0)
		{
			num2 *= -1;
			return CreateResult(System.Math.Round(num / System.Math.Pow(10.0, num2), 0, MidpointRounding.AwayFromZero) * System.Math.Pow(10.0, num2), DataType.Integer);
		}
		return CreateResult(System.Math.Round(num, num2, MidpointRounding.AwayFromZero), DataType.Decimal);
	}
}
