using System;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class Power : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 2);
		double x = ArgToDecimal(arguments, 0);
		double y = ArgToDecimal(arguments, 1);
		double num = System.Math.Pow(x, y);
		return CreateResult(num, DataType.Decimal);
	}
}
