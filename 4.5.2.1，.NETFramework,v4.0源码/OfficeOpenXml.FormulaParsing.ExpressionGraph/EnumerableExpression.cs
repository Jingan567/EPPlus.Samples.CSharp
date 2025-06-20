using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph;

public class EnumerableExpression : Expression
{
	private readonly IExpressionCompiler _expressionCompiler;

	public override bool IsGroupedExpression => false;

	public EnumerableExpression()
		: this(new ExpressionCompiler())
	{
	}

	public EnumerableExpression(IExpressionCompiler expressionCompiler)
	{
		_expressionCompiler = expressionCompiler;
	}

	public override Expression PrepareForNextChild()
	{
		return this;
	}

	public override CompileResult Compile()
	{
		List<object> list = new List<object>();
		foreach (Expression child in base.Children)
		{
			list.Add(_expressionCompiler.Compile(new List<Expression> { child }).Result);
		}
		return new CompileResult(list, DataType.Enumerable);
	}
}
