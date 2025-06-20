namespace OfficeOpenXml.FormulaParsing.ExpressionGraph;

public class FunctionArgumentExpression : GroupExpression
{
	private readonly Expression _function;

	public override bool ParentIsLookupFunction
	{
		get
		{
			return base.ParentIsLookupFunction;
		}
		set
		{
			base.ParentIsLookupFunction = value;
			foreach (Expression child in base.Children)
			{
				child.ParentIsLookupFunction = value;
			}
		}
	}

	public override bool IsGroupedExpression => false;

	public FunctionArgumentExpression(Expression function)
		: base(isNegated: false)
	{
		_function = function;
	}

	public override Expression PrepareForNextChild()
	{
		return _function.PrepareForNextChild();
	}
}
