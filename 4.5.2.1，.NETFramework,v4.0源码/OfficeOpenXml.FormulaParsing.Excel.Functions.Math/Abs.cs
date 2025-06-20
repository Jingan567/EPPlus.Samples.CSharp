using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class Abs : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		FunctionArgument[] array = (arguments as FunctionArgument[]) ?? arguments.ToArray();
		ValidateArguments(array, 1);
		if (array.ElementAt(0).Value == null)
		{
			return CreateResult(0.0, DataType.Decimal);
		}
		double num = ArgToDecimal(array, 0);
		if (num < 0.0)
		{
			num *= -1.0;
		}
		return CreateResult(num, DataType.Decimal);
	}
}
