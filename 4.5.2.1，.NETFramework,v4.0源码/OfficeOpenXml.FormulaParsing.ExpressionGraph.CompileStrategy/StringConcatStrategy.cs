namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.CompileStrategy;

public class StringConcatStrategy : CompileStrategy
{
	public StringConcatStrategy(Expression expression)
		: base(expression)
	{
	}

	public override Expression Compile()
	{
		Expression expression = ((_expression is ExcelAddressExpression) ? _expression : ExpressionConverter.Instance.ToStringExpression(_expression));
		expression.Prev = _expression.Prev;
		expression.Next = _expression.Next;
		if (_expression.Prev != null)
		{
			_expression.Prev.Next = expression;
		}
		if (_expression.Next != null)
		{
			_expression.Next.Prev = expression;
		}
		return expression.MergeWithNext();
	}
}
