using System;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class Ceiling : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 2);
		double num = ArgToDecimal(arguments, 0);
		double num2 = ArgToDecimal(arguments, 1);
		ValidateNumberAndSign(num, num2);
		if (num2 < 1.0 && num2 > 0.0)
		{
			double num3 = System.Math.Floor(num);
			int num4 = (int)((num - num3) / num2) + 1;
			return CreateResult(num3 + (double)num4 * num2, DataType.Decimal);
		}
		if (num2 == 1.0)
		{
			return CreateResult(System.Math.Ceiling(num), DataType.Decimal);
		}
		if (num % num2 == 0.0)
		{
			return CreateResult(num, DataType.Decimal);
		}
		double num5 = num - num % num2 + num2;
		return CreateResult(num5, DataType.Decimal);
	}

	private void ValidateNumberAndSign(double number, double sign)
	{
		if (number > 0.0 && sign < 0.0)
		{
			string text = $"num: {number}, sign: {sign}";
			throw new InvalidOperationException("Ceiling cannot handle a negative significance when the number is positive" + text);
		}
	}
}
