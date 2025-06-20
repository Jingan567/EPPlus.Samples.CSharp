using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

public class Exact : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 2);
		object valueFirst = arguments.ElementAt(0).ValueFirst;
		object valueFirst2 = arguments.ElementAt(1).ValueFirst;
		if (valueFirst == null && valueFirst2 == null)
		{
			return CreateResult(true, DataType.Boolean);
		}
		if ((valueFirst == null && valueFirst2 != null) || (valueFirst != null && valueFirst2 == null))
		{
			return CreateResult(false, DataType.Boolean);
		}
		int num = string.Compare(valueFirst.ToString(), valueFirst2.ToString(), StringComparison.Ordinal);
		return CreateResult(num == 0, DataType.Boolean);
	}
}
