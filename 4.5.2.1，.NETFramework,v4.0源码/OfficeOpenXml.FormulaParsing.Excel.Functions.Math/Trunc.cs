using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class Trunc : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		double d = ArgToDecimal(arguments, 0);
		if (arguments.Count() == 1)
		{
			return CreateResult(System.Math.Truncate(d), DataType.Decimal);
		}
		ArgToInt(arguments, 1);
		return context.Configuration.FunctionRepository.GetFunction("rounddown").Execute(arguments, context);
	}
}
