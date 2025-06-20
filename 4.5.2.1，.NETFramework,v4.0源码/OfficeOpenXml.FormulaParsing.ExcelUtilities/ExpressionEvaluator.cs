using System;
using System.Text.RegularExpressions;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities;

public class ExpressionEvaluator
{
	private readonly WildCardValueMatcher _wildCardValueMatcher;

	private readonly CompileResultFactory _compileResultFactory;

	public ExpressionEvaluator()
		: this(new WildCardValueMatcher(), new CompileResultFactory())
	{
	}

	public ExpressionEvaluator(WildCardValueMatcher wildCardValueMatcher, CompileResultFactory compileResultFactory)
	{
		_wildCardValueMatcher = wildCardValueMatcher;
		_compileResultFactory = compileResultFactory;
	}

	private string GetNonAlphanumericStartChars(string expression)
	{
		if (!string.IsNullOrEmpty(expression))
		{
			if (Regex.IsMatch(expression, "^([^a-zA-Z0-9]{2})"))
			{
				return expression.Substring(0, 2);
			}
			if (Regex.IsMatch(expression, "^([^a-zA-Z0-9]{1})"))
			{
				return expression.Substring(0, 1);
			}
		}
		return null;
	}

	private bool EvaluateOperator(object left, object right, IOperator op)
	{
		CompileResult left2 = _compileResultFactory.Create(left);
		CompileResult right2 = _compileResultFactory.Create(right);
		CompileResult compileResult = op.Apply(left2, right2);
		if (compileResult.DataType != DataType.Boolean)
		{
			throw new ArgumentException("Illegal operator in expression");
		}
		return (bool)compileResult.Result;
	}

	public bool TryConvertToDouble(object op, out double d)
	{
		if (op is double || op is int)
		{
			d = Convert.ToDouble(op);
			return true;
		}
		if (op is DateTime)
		{
			d = ((DateTime)op).ToOADate();
			return true;
		}
		if (op != null && double.TryParse(op.ToString(), out d))
		{
			return true;
		}
		d = 0.0;
		return false;
	}

	public bool Evaluate(object left, string expression)
	{
		if (expression == string.Empty)
		{
			return left == null;
		}
		string nonAlphanumericStartChars = GetNonAlphanumericStartChars(expression);
		if (!string.IsNullOrEmpty(nonAlphanumericStartChars) && nonAlphanumericStartChars != "-" && OperatorsDict.Instance.TryGetValue(nonAlphanumericStartChars, out var value))
		{
			string text = expression.Replace(nonAlphanumericStartChars, string.Empty);
			if (left == null && text == string.Empty)
			{
				return value.Operator == Operators.Equals;
			}
			if ((left == null) ^ (text == string.Empty))
			{
				return value.Operator == Operators.NotEqualTo;
			}
			double d;
			bool flag = TryConvertToDouble(left, out d);
			double result;
			bool flag2 = double.TryParse(text, out result);
			DateTime result2;
			bool flag3 = DateTime.TryParse(text, out result2);
			if (flag && flag2)
			{
				return EvaluateOperator(d, result, value);
			}
			if (flag && flag3)
			{
				return EvaluateOperator(d, result2.ToOADate(), value);
			}
			if (flag != flag2)
			{
				return value.Operator == Operators.NotEqualTo;
			}
			return EvaluateOperator(left, text, value);
		}
		return _wildCardValueMatcher.IsMatch(expression, left) == 0;
	}
}
