using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;

public class Or : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		for (int i = 0; i < arguments.Count(); i++)
		{
			if (ArgToBool(arguments, i))
			{
				return new CompileResult(true, DataType.Boolean);
			}
		}
		return new CompileResult(false, DataType.Boolean);
	}
}
