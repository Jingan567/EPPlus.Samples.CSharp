using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

public class Left : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 2);
		string text = ArgToString(arguments, 0);
		int num = ArgToInt(arguments, 1);
		if (text.Length < num)
		{
			num = text.Length;
		}
		return CreateResult(text.Substring(0, num), DataType.String);
	}
}
