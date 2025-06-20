using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;

public class Day : DateParsingFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		object firstValue = GetFirstValue(arguments);
		return CreateResult(ParseDate(arguments, firstValue).Day, DataType.Integer);
	}
}
