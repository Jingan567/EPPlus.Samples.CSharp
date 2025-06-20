using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Information;

public class IsNumber : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		object firstValue = GetFirstValue(arguments);
		return CreateResult(IsNumeric(firstValue), DataType.Boolean);
	}
}
