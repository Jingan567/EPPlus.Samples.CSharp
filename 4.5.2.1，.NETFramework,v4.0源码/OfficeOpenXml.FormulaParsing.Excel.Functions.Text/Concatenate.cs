using System.Collections.Generic;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

public class Concatenate : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		if (arguments == null)
		{
			return CreateResult(string.Empty, DataType.String);
		}
		StringBuilder stringBuilder = new StringBuilder();
		foreach (FunctionArgument argument in arguments)
		{
			object valueFirst = argument.ValueFirst;
			if (valueFirst != null)
			{
				stringBuilder.Append(valueFirst);
			}
		}
		return CreateResult(stringBuilder.ToString(), DataType.String);
	}
}
