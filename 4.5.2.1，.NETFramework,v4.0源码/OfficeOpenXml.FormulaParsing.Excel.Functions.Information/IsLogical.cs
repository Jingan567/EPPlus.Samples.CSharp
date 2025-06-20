using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Information;

public class IsLogical : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		FunctionArgument[] arguments2 = (arguments as FunctionArgument[]) ?? arguments.ToArray();
		ValidateArguments(arguments2, 1);
		object firstValue = GetFirstValue(arguments);
		return CreateResult(firstValue is bool, DataType.Boolean);
	}
}
