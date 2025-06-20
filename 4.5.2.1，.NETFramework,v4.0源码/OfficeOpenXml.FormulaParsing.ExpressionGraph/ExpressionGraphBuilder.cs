using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph;

public class ExpressionGraphBuilder : IExpressionGraphBuilder
{
	private readonly ExpressionGraph _graph = new ExpressionGraph();

	private readonly IExpressionFactory _expressionFactory;

	private readonly ParsingContext _parsingContext;

	private int _tokenIndex;

	private bool _negateNextExpression;

	public ExpressionGraphBuilder(ExcelDataProvider excelDataProvider, ParsingContext parsingContext)
		: this(new ExpressionFactory(excelDataProvider, parsingContext), parsingContext)
	{
	}

	public ExpressionGraphBuilder(IExpressionFactory expressionFactory, ParsingContext parsingContext)
	{
		_expressionFactory = expressionFactory;
		_parsingContext = parsingContext;
	}

	public ExpressionGraph Build(IEnumerable<Token> tokens)
	{
		_tokenIndex = 0;
		_graph.Reset();
		Token[] tokens2 = ((tokens != null) ? tokens.ToArray() : new Token[0]);
		BuildUp(tokens2, null);
		return _graph;
	}

	private void BuildUp(Token[] tokens, Expression parent)
	{
		while (_tokenIndex < tokens.Length)
		{
			Token token = tokens[_tokenIndex];
			IOperator value = null;
			if (token.TokenType == TokenType.Operator && OperatorsDict.Instance.TryGetValue(token.Value, out value))
			{
				SetOperatorOnExpression(parent, value);
			}
			else if (token.TokenType == TokenType.Function)
			{
				BuildFunctionExpression(tokens, parent, token.Value);
			}
			else if (token.TokenType == TokenType.OpeningEnumerable)
			{
				_tokenIndex++;
				BuildEnumerableExpression(tokens, parent);
			}
			else if (token.TokenType == TokenType.OpeningParenthesis)
			{
				_tokenIndex++;
				BuildGroupExpression(tokens, parent);
			}
			else
			{
				if (token.TokenType == TokenType.ClosingParenthesis || token.TokenType == TokenType.ClosingEnumerable)
				{
					break;
				}
				if (token.TokenType == TokenType.WorksheetName)
				{
					StringBuilder stringBuilder = new StringBuilder();
					stringBuilder.Append(tokens[_tokenIndex++].Value);
					stringBuilder.Append(tokens[_tokenIndex++].Value);
					stringBuilder.Append(tokens[_tokenIndex++].Value);
					stringBuilder.Append(tokens[_tokenIndex].Value);
					Token token2 = new Token(stringBuilder.ToString(), TokenType.ExcelAddress);
					CreateAndAppendExpression(ref parent, token2);
				}
				else if (token.TokenType == TokenType.Negator)
				{
					_negateNextExpression = true;
				}
				else if (token.TokenType == TokenType.Percent)
				{
					SetOperatorOnExpression(parent, Operator.Percent);
					if (parent == null)
					{
						_graph.Add(ConstantExpressions.Percent);
					}
					else
					{
						parent.AddChild(ConstantExpressions.Percent);
					}
				}
				else
				{
					CreateAndAppendExpression(ref parent, token);
				}
			}
			_tokenIndex++;
		}
	}

	private void BuildEnumerableExpression(Token[] tokens, Expression parent)
	{
		if (parent == null)
		{
			_graph.Add(new EnumerableExpression());
			BuildUp(tokens, _graph.Current);
		}
		else
		{
			EnumerableExpression enumerableExpression = new EnumerableExpression();
			parent.AddChild(enumerableExpression);
			BuildUp(tokens, enumerableExpression);
		}
	}

	private void CreateAndAppendExpression(ref Expression parent, Token token)
	{
		if (IsWaste(token))
		{
			return;
		}
		if (parent != null && (token.TokenType == TokenType.Comma || token.TokenType == TokenType.SemiColon))
		{
			parent = parent.PrepareForNextChild();
			return;
		}
		if (_negateNextExpression)
		{
			token.Negate();
			_negateNextExpression = false;
		}
		Expression expression = _expressionFactory.Create(token);
		if (parent == null)
		{
			_graph.Add(expression);
		}
		else
		{
			parent.AddChild(expression);
		}
	}

	private bool IsWaste(Token token)
	{
		if (token.TokenType == TokenType.String)
		{
			return true;
		}
		return false;
	}

	private void BuildFunctionExpression(Token[] tokens, Expression parent, string funcName)
	{
		if (parent == null)
		{
			_graph.Add(new FunctionExpression(funcName, _parsingContext, _negateNextExpression));
			_negateNextExpression = false;
			HandleFunctionArguments(tokens, _graph.Current);
		}
		else
		{
			FunctionExpression functionExpression = new FunctionExpression(funcName, _parsingContext, _negateNextExpression);
			_negateNextExpression = false;
			parent.AddChild(functionExpression);
			HandleFunctionArguments(tokens, functionExpression);
		}
	}

	private void HandleFunctionArguments(Token[] tokens, Expression function)
	{
		_tokenIndex++;
		if (tokens.ElementAt(_tokenIndex).TokenType != TokenType.OpeningParenthesis)
		{
			throw new ExcelErrorValueException(eErrorType.Value);
		}
		_tokenIndex++;
		BuildUp(tokens, function.Children.First());
	}

	private void BuildGroupExpression(Token[] tokens, Expression parent)
	{
		if (parent == null)
		{
			_graph.Add(new GroupExpression(_negateNextExpression));
			_negateNextExpression = false;
			BuildUp(tokens, _graph.Current);
			return;
		}
		if (parent.IsGroupedExpression || parent is FunctionArgumentExpression)
		{
			GroupExpression groupExpression = new GroupExpression(_negateNextExpression);
			_negateNextExpression = false;
			parent.AddChild(groupExpression);
			BuildUp(tokens, groupExpression);
		}
		BuildUp(tokens, parent);
	}

	private void SetOperatorOnExpression(Expression parent, IOperator op)
	{
		if (parent == null)
		{
			_graph.Current.Operator = op;
			return;
		}
		Expression expression;
		if (parent is FunctionArgumentExpression)
		{
			expression = parent.Children.Last();
		}
		else
		{
			expression = parent.Children.Last();
			if (expression is FunctionArgumentExpression)
			{
				expression = expression.Children.Last();
			}
		}
		expression.Operator = op;
	}
}
