using System;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class Rounddown : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 2);
		double num = ArgToDecimal(arguments, 0);
		int num2 = ArgToInt(arguments, 1);
		int num3 = ((!(num < 0.0)) ? 1 : (-1));
		num *= (double)num3;
		double num4;
		if (num2 > 0)
		{
			num4 = RoundDownDecimalNumber(num, num2);
		}
		else
		{
			num4 = (int)System.Math.Floor(num);
			num4 -= num4 % System.Math.Pow(10.0, num2 * -1);
		}
		return CreateResult(num4 * (double)num3, DataType.Decimal);
	}

	private static double RoundDownDecimalNumber(double number, int nDecimals)
	{
		double num = System.Math.Floor(number);
		double num2 = number - num;
		num2 = System.Math.Pow(10.0, nDecimals) * num2;
		num2 = System.Math.Truncate(num2) / System.Math.Pow(10.0, nDecimals);
		return num + num2;
	}
}
