using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

public class T : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		object valueFirst = arguments.ElementAt(0).ValueFirst;
		if (valueFirst is string)
		{
			return CreateResult(valueFirst, DataType.String);
		}
		return CreateResult(string.Empty, DataType.String);
	}
}
