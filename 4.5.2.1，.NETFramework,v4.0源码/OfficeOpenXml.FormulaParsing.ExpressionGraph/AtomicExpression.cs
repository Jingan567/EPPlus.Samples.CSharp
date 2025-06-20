namespace OfficeOpenXml.FormulaParsing.ExpressionGraph;

public abstract class AtomicExpression : Expression
{
	public override bool IsGroupedExpression => false;

	public AtomicExpression(string expression)
		: base(expression)
	{
	}
}
