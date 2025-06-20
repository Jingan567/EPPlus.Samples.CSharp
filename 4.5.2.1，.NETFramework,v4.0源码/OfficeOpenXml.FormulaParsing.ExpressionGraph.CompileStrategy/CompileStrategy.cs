namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.CompileStrategy;

public abstract class CompileStrategy
{
	protected readonly Expression _expression;

	public CompileStrategy(Expression expression)
	{
		_expression = expression;
	}

	public abstract Expression Compile();
}
