namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.CompileStrategy;

public class DefaultCompileStrategy : CompileStrategy
{
	public DefaultCompileStrategy(Expression expression)
		: base(expression)
	{
	}

	public override Expression Compile()
	{
		return _expression.MergeWithNext();
	}
}
