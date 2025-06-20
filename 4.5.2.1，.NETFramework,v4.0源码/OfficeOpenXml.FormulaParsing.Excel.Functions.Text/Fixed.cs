using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

public class Fixed : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		double num = ArgToDecimal(arguments, 0);
		int num2 = 2;
		bool flag = false;
		if (arguments.Count() > 1)
		{
			num2 = ArgToInt(arguments, 1);
		}
		if (arguments.Count() > 2)
		{
			flag = ArgToBool(arguments, 2);
		}
		string text = (flag ? "F" : "N") + num2.ToString(CultureInfo.InvariantCulture);
		if (num2 < 0)
		{
			num -= num % System.Math.Pow(10.0, num2 * -1);
			num = System.Math.Floor(num);
			text = (flag ? "F0" : "N0");
		}
		string result = num.ToString(text);
		return CreateResult(result, DataType.String);
	}
}
