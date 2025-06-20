using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Database;

public class Dvar : DatabaseFunction
{
	public Dvar()
		: this(new RowMatcher())
	{
	}

	public Dvar(RowMatcher rowMatcher)
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
		return CreateResult(VarMethods.Var(matchingValues), DataType.Integer);
	}
}
