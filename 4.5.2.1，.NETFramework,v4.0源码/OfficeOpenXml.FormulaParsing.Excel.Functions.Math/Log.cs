using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class Log : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		double a = ArgToDecimal(arguments, 0);
		if (arguments.Count() == 1)
		{
			return CreateResult(System.Math.Log(a, 10.0), DataType.Decimal);
		}
		double newBase = ArgToDecimal(arguments, 1);
		return CreateResult(System.Math.Log(a, newBase), DataType.Decimal);
	}
}
