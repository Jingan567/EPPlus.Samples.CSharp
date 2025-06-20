using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

public class Right : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 2);
		string text = ArgToString(arguments, 0);
		int num = ArgToInt(arguments, 1);
		int num2 = text.Length - num;
		if (num2 < 0)
		{
			num2 = 0;
		}
		return CreateResult(text.Substring(num2, text.Length - num2), DataType.String);
	}
}
