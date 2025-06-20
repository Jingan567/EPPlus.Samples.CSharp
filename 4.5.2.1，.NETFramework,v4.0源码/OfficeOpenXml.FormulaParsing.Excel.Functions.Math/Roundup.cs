using System;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class Roundup : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 2);
		double num = ArgToDecimal(arguments, 0);
		int num2 = ArgToInt(arguments, 1);
		double num3 = ((num >= 0.0) ? (System.Math.Ceiling(num * System.Math.Pow(10.0, num2)) / System.Math.Pow(10.0, num2)) : (System.Math.Floor(num * System.Math.Pow(10.0, num2)) / System.Math.Pow(10.0, num2)));
		return CreateResult(num3, DataType.Decimal);
	}
}
