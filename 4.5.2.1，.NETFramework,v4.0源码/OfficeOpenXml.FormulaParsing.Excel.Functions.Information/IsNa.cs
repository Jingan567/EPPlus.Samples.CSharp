using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Information;

public class IsNa : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		if (arguments == null || arguments.Count() == 0)
		{
			return CreateResult(false, DataType.Boolean);
		}
		object firstValue = GetFirstValue(arguments);
		if (firstValue is ExcelErrorValue && ((ExcelErrorValue)firstValue).Type == eErrorType.NA)
		{
			return CreateResult(true, DataType.Boolean);
		}
		return CreateResult(false, DataType.Boolean);
	}
}
