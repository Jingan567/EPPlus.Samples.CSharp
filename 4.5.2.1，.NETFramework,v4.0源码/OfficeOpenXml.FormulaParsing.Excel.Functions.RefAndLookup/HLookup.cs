using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

public class HLookup : LookupFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 3);
		LookupArguments lookupArgs = new LookupArguments(arguments, context);
		ThrowExcelErrorValueExceptionIf(() => lookupArgs.LookupIndex < 1, eErrorType.Value);
		LookupNavigator navigator = LookupNavigatorFactory.Create(LookupDirection.Horizontal, lookupArgs, context);
		return Lookup(navigator, lookupArgs);
	}
}
