using System;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class RandBetween : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 2);
		double num = ArgToDecimal(arguments, 0);
		double high = ArgToDecimal(arguments, 1);
		object result = new Rand().Execute(new FunctionArgument[0], context).Result;
		double d = CalulateDiff(high, num) * (double)result + 1.0;
		d = System.Math.Floor(d);
		return CreateResult(num + d, DataType.Integer);
	}

	private double CalulateDiff(double high, double low)
	{
		if (high > 0.0 && low < 0.0)
		{
			return high + low * -1.0;
		}
		if (high < 0.0 && low < 0.0)
		{
			return high * -1.0 - low * -1.0;
		}
		return high - low;
	}
}
