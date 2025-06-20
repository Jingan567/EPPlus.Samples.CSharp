using System;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class Atan2 : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 2);
		double x = ArgToDecimal(arguments, 0);
		double y = ArgToDecimal(arguments, 1);
		return CreateResult(System.Math.Atan2(y, x), DataType.Decimal);
	}
}
