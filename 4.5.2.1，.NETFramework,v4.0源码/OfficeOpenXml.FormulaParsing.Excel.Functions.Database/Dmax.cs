using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Database;

public class Dmax : DatabaseFunction
{
	public Dmax()
		: this(new RowMatcher())
	{
	}

	public Dmax(RowMatcher rowMatcher)
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
		return CreateResult(matchingValues.Max(), DataType.Integer);
	}
}
