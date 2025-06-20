using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;

public class If : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 3);
		bool num = ArgToBool(arguments, 0);
		object value = arguments.ElementAt(1).Value;
		object value2 = arguments.ElementAt(2).Value;
		CompileResultFactory compileResultFactory = new CompileResultFactory();
		if (!num)
		{
			return compileResultFactory.Create(value2);
		}
		return compileResultFactory.Create(value);
	}
}
