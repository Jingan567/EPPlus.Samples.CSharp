using OfficeOpenXml.FormulaParsing.Excel.Operators;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.CompileStrategy;

public class CompileStrategyFactory : ICompileStrategyFactory
{
	public CompileStrategy Create(Expression expression)
	{
		if (expression.Operator.Operator == Operators.Concat)
		{
			return new StringConcatStrategy(expression);
		}
		return new DefaultCompileStrategy(expression);
	}
}
