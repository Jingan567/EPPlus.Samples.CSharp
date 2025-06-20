using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Information;

public class IsText : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		if (arguments.Count() == 1 && arguments.ElementAt(0).Value != null)
		{
			return CreateResult(GetFirstValue(arguments) is string, DataType.Boolean);
		}
		return CreateResult(false, DataType.Boolean);
	}
}
