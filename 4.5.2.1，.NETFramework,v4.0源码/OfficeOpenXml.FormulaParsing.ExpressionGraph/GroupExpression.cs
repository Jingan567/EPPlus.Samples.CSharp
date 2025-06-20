namespace OfficeOpenXml.FormulaParsing.ExpressionGraph;

public class GroupExpression : Expression
{
	private readonly IExpressionCompiler _expressionCompiler;

	private readonly bool _isNegated;

	public override bool IsGroupedExpression => true;

	public GroupExpression(bool isNegated)
		: this(isNegated, new ExpressionCompiler())
	{
	}

	public GroupExpression(bool isNegated, IExpressionCompiler expressionCompiler)
	{
		_expressionCompiler = expressionCompiler;
		_isNegated = isNegated;
	}

	public override CompileResult Compile()
	{
		CompileResult compileResult = _expressionCompiler.Compile(base.Children);
		if (compileResult.IsNumeric && _isNegated)
		{
			return new CompileResult(compileResult.ResultNumeric * -1.0, compileResult.DataType);
		}
		return compileResult;
	}
}
