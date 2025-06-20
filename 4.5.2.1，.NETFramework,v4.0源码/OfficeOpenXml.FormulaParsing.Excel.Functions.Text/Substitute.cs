using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

public class Substitute : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 3);
		string text = ArgToString(arguments, 0);
		string oldValue = ArgToString(arguments, 1);
		string newValue = ArgToString(arguments, 2);
		string result = text.Replace(oldValue, newValue);
		return CreateResult(result, DataType.String);
	}
}
