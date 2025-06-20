using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Database;

public class Dmin : DatabaseFunction
{
	public Dmin()
		: this(new RowMatcher())
	{
	}

	public Dmin(RowMatcher rowMatcher)
		: base(rowMatcher)
	{
	}

	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 3);
		IEnumerable<double> matchingValues = GetMatchingValues(arguments, context);
		if (!matchingValues.Any())
		{
			return CreateResult(0.0, DataType.Integer);
		}
		return CreateResult(matchingValues.Min(), DataType.Integer);
	}
}
