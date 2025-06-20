using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Excel.Operators;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph;

public abstract class Expression
{
	private readonly List<Expression> _children = new List<Expression>();

	protected string ExpressionString { get; private set; }

	public IEnumerable<Expression> Children => _children;

	public Expression Next { get; set; }

	public Expression Prev { get; set; }

	public IOperator Operator { get; set; }

	public abstract bool IsGroupedExpression { get; }

	public virtual bool ParentIsLookupFunction { get; set; }

	public virtual bool HasChildren => _children.Any();

	public Expression()
	{
	}

	public Expression(string expression)
	{
		ExpressionString = expression;
		Operator = null;
	}

	public virtual Expression PrepareForNextChild()
	{
		return this;
	}

	public virtual Expression AddChild(Expression child)
	{
		if (_children.Any())
		{
			Expression expression2 = (child.Prev = _children.Last());
			expression2.Next = child;
		}
		_children.Add(child);
		return child;
	}

	public virtual Expression MergeWithNext()
	{
		Expression expression = this;
		if (Next != null && Operator != null)
		{
			CompileResult compileResult = Operator.Apply(Compile(), Next.Compile());
			expression = ExpressionConverter.Instance.FromCompileResult(compileResult);
			if (expression is ExcelErrorExpression)
			{
				expression.Next = null;
				expression.Prev = null;
				return expression;
			}
			if (Next != null)
			{
				expression.Operator = Next.Operator;
			}
			else
			{
				expression.Operator = null;
			}
			expression.Next = Next.Next;
			if (expression.Next != null)
			{
				expression.Next.Prev = expression;
			}
			expression.Prev = Prev;
			if (Prev != null)
			{
				Prev.Next = expression;
			}
			return expression;
		}
		throw new FormatException("Invalid formula syntax. Operator missing expression.");
	}

	public abstract CompileResult Compile();
}
