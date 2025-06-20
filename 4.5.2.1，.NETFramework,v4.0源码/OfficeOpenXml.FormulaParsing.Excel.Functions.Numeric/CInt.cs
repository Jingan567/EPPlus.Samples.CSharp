using System;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Numeric;

public class CInt : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		double d = ArgToDecimal(arguments, 0);
		return CreateResult((int)System.Math.Floor(d), DataType.Integer);
	}
}
