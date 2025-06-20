using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Information;

public class IsNonText : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		FunctionArgument functionArgument = arguments.ElementAt(0);
		if (functionArgument.Value == null || functionArgument.ValueIsExcelError)
		{
			return CreateResult(false, DataType.Boolean);
		}
		return CreateResult(!(functionArgument.Value is string), DataType.Boolean);
	}
}
