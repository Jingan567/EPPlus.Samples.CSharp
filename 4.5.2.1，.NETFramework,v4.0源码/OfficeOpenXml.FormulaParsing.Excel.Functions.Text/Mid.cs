using System;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

public class Mid : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 3);
		string text = ArgToString(arguments, 0);
		int num = ArgToInt(arguments, 1);
		int num2 = ArgToInt(arguments, 2);
		if (num <= 0)
		{
			throw new ArgumentException("Argument start can't be less than 1");
		}
		if (num > text.Length)
		{
			return CreateResult("", DataType.String);
		}
		string result = text.Substring(num - 1, (num - 1 + num2 < text.Length) ? num2 : (text.Length - num + 1));
		return CreateResult(result, DataType.String);
	}
}
