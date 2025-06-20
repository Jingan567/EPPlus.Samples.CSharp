using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.CompileStrategy;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph;

public class ExpressionCompiler : IExpressionCompiler
{
	private IEnumerable<Expression> _expressions;

	private IExpressionConverter _expressionConverter;

	private ICompileStrategyFactory _compileStrategyFactory;

	public ExpressionCompiler()
		: this(new ExpressionConverter(), new CompileStrategyFactory())
	{
	}

	public ExpressionCompiler(IExpressionConverter expressionConverter, ICompileStrategyFactory compileStrategyFactory)
	{
		_expressionConverter = expressionConverter;
		_compileStrategyFactory = compileStrategyFactory;
	}

	public CompileResult Compile(IEnumerable<Expression> expressions)
	{
		_expressions = expressions;
		return PerformCompilation();
	}

	public CompileResult Compile(string worksheet, int row, int column, IEnumerable<Expression> expressions)
	{
		_expressions = expressions;
		return PerformCompilation(worksheet, row, column);
	}

	private CompileResult PerformCompilation(string worksheet = "", int row = -1, int column = -1)
	{
		IEnumerable<Expression> source = HandleGroupedExpressions();
		while (source.Any((Expression x) => x.Operator != null))
		{
			int precedence = FindLowestPrecedence();
			source = HandlePrecedenceLevel(precedence);
		}
		if (_expressions.Any())
		{
			return source.First().Compile();
		}
		return CompileResult.Empty;
	}

	private IEnumerable<Expression> HandleGroupedExpressions()
	{
		if (!_expressions.Any())
		{
			return Enumerable.Empty<Expression>();
		}
		Expression expression = _expressions.First();
		foreach (Expression item in _expressions.Where((Expression x) => x.IsGroupedExpression))
		{
			CompileResult compileResult = item.Compile();
			if (compileResult != CompileResult.Empty)
			{
				Expression expression2 = _expressionConverter.FromCompileResult(compileResult);
				expression2.Operator = item.Operator;
				expression2.Prev = item.Prev;
				expression2.Next = item.Next;
				if (item.Prev != null)
				{
					item.Prev.Next = expression2;
				}
				if (item.Next != null)
				{
					item.Next.Prev = expression2;
				}
				if (item == expression)
				{
					expression = expression2;
				}
			}
		}
		return RefreshList(expression);
	}

	private IEnumerable<Expression> HandlePrecedenceLevel(int precedence)
	{
		Expression expression = _expressions.First();
		IEnumerable<Expression> source = _expressions.Where((Expression x) => x.Operator != null && x.Operator.Precedence == precedence);
		source.Last();
		Expression expression2 = source.First();
		do
		{
			Expression expression3 = _compileStrategyFactory.Create(expression2).Compile();
			if (expression3 is ExcelErrorExpression)
			{
				return RefreshList(expression3);
			}
			if (expression2 == expression)
			{
				expression = expression3;
			}
			expression2 = expression3;
		}
		while (expression2 != null && expression2.Operator != null && expression2.Operator.Precedence == precedence);
		return RefreshList(expression);
	}

	private int FindLowestPrecedence()
	{
		return _expressions.Where((Expression x) => x.Operator != null).Min((Expression x) => x.Operator.Precedence);
	}

	private IEnumerable<Expression> RefreshList(Expression first)
	{
		List<Expression> list = new List<Expression>();
		Expression expression = first;
		list.Add(expression);
		while (expression.Next != null)
		{
			list.Add(expression.Next);
			expression = expression.Next;
		}
		_expressions = list;
		return list;
	}
}
