using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;

public class Second : DateParsingFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		object value = arguments.ElementAt(0).Value;
		return CreateResult(ParseDate(arguments, value).Second, DataType.Integer);
	}
}
