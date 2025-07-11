using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

public class CharFunction : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		int number = ArgToInt(arguments, 0);
		ThrowExcelErrorValueExceptionIf(() => number < 1 || number > 255, eErrorType.Value);
		return CreateResult(((char)number).ToString(), DataType.String);
	}
}
