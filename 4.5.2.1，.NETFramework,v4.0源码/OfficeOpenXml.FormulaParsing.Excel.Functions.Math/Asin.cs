using System;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class Asin : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		double d = ArgToDecimal(arguments, 0);
		return CreateResult(System.Math.Asin(d), DataType.Decimal);
	}
}
