using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class VarP : HiddenValuesHandlingFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		return new CompileResult(VarMethods.VarP(ArgsToDoubleEnumerable(base.IgnoreHiddenValues, ignoreErrors: false, arguments, context)), DataType.Decimal);
	}
}
