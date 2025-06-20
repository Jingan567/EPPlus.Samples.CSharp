using System.Collections.Generic;
using System.Diagnostics;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

public class VLookup : LookupFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		Stopwatch stopwatch = null;
		if (context.Debug)
		{
			stopwatch = new Stopwatch();
			stopwatch.Start();
		}
		ValidateArguments(arguments, 3);
		LookupArguments lookupArguments = new LookupArguments(arguments, context);
		LookupNavigator navigator = LookupNavigatorFactory.Create(LookupDirection.Vertical, lookupArguments, context);
		CompileResult result = Lookup(navigator, lookupArguments);
		if (context.Debug)
		{
			stopwatch.Stop();
			context.Configuration.Logger.LogFunction("VLOOKUP", stopwatch.ElapsedMilliseconds);
		}
		return result;
	}
}
