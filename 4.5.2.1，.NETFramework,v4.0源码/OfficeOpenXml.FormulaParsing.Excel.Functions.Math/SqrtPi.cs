using System;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class SqrtPi : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		double num = ArgToDecimal(arguments, 0);
		return CreateResult(System.Math.Sqrt(num * System.Math.PI), DataType.Decimal);
	}
}
