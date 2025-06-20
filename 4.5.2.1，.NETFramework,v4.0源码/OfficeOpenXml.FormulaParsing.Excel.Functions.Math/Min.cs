using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class Min : HiddenValuesHandlingFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		IEnumerable<ExcelDoubleCellValue> source = ArgsToDoubleEnumerable(base.IgnoreHiddenValues, ignoreErrors: false, arguments, context);
		return CreateResult(source.Min(), DataType.Decimal);
	}
}
