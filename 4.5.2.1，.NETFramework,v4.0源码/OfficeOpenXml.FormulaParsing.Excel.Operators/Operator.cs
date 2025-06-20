using System;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Operators;

public class Operator : IOperator
{
	private const int PrecedencePercent = 2;

	private const int PrecedenceExp = 4;

	private const int PrecedenceMultiplyDevide = 6;

	private const int PrecedenceIntegerDivision = 8;

	private const int PrecedenceModulus = 10;

	private const int PrecedenceAddSubtract = 12;

	private const int PrecedenceConcat = 15;

	private const int PrecedenceComparison = 25;

	private readonly Func<CompileResult, CompileResult, CompileResult> _implementation;

	private readonly int _precedence;

	private readonly Operators _operator;

	private static IOperator _plus;

	private static IOperator _minus;

	private static IOperator _multiply;

	private static IOperator _divide;

	private static IOperator _greaterThan;

	private static IOperator _eq;

	private static IOperator _notEqualsTo;

	private static IOperator _greaterThanOrEqual;

	private static IOperator _lessThan;

	private static IOperator _percent;

	int IOperator.Precedence => _precedence;

	Operators IOperator.Operator => _operator;

	public static IOperator Plus => _plus ?? (_plus = new Operator(Operators.Plus, 12, delegate(CompileResult l, CompileResult r)
	{
		l = ((l == null || l.Result == null) ? new CompileResult(0, DataType.Integer) : l);
		r = ((r == null || r.Result == null) ? new CompileResult(0, DataType.Integer) : r);
		if (EitherIsError(l, r, out var errorVal))
		{
			return new CompileResult(errorVal);
		}
		if (l.DataType == DataType.Integer && r.DataType == DataType.Integer)
		{
			return new CompileResult(l.ResultNumeric + r.ResultNumeric, DataType.Integer);
		}
		return ((l.IsNumeric || l.IsNumericString || l.IsDateString || l.Result is ExcelDataProvider.IRangeInfo) && (r.IsNumeric || r.IsNumericString || r.IsDateString || r.Result is ExcelDataProvider.IRangeInfo)) ? new CompileResult(l.ResultNumeric + r.ResultNumeric, DataType.Decimal) : new CompileResult(eErrorType.Value);
	}));

	public static IOperator Minus => _minus ?? (_minus = new Operator(Operators.Minus, 12, delegate(CompileResult l, CompileResult r)
	{
		l = ((l == null || l.Result == null) ? new CompileResult(0, DataType.Integer) : l);
		r = ((r == null || r.Result == null) ? new CompileResult(0, DataType.Integer) : r);
		if (l.DataType == DataType.Integer && r.DataType == DataType.Integer)
		{
			return new CompileResult(l.ResultNumeric - r.ResultNumeric, DataType.Integer);
		}
		return ((l.IsNumeric || l.IsNumericString || l.IsDateString || l.Result is ExcelDataProvider.IRangeInfo) && (r.IsNumeric || r.IsNumericString || r.IsDateString || r.Result is ExcelDataProvider.IRangeInfo)) ? new CompileResult(l.ResultNumeric - r.ResultNumeric, DataType.Decimal) : new CompileResult(eErrorType.Value);
	}));

	public static IOperator Multiply => _multiply ?? (_multiply = new Operator(Operators.Multiply, 6, delegate(CompileResult l, CompileResult r)
	{
		l = l ?? new CompileResult(0, DataType.Integer);
		r = r ?? new CompileResult(0, DataType.Integer);
		if (l.DataType == DataType.Integer && r.DataType == DataType.Integer)
		{
			return new CompileResult(l.ResultNumeric * r.ResultNumeric, DataType.Integer);
		}
		return ((l.IsNumeric || l.IsNumericString || l.IsDateString || l.Result is ExcelDataProvider.IRangeInfo) && (r.IsNumeric || r.IsNumericString || r.IsDateString || r.Result is ExcelDataProvider.IRangeInfo)) ? new CompileResult(l.ResultNumeric * r.ResultNumeric, DataType.Decimal) : new CompileResult(eErrorType.Value);
	}));

	public static IOperator Divide => _divide ?? (_divide = new Operator(Operators.Divide, 6, delegate(CompileResult l, CompileResult r)
	{
		if ((!l.IsNumeric && !l.IsNumericString && !l.IsDateString && !(l.Result is ExcelDataProvider.IRangeInfo)) || (!r.IsNumeric && !r.IsNumericString && !r.IsDateString && !(r.Result is ExcelDataProvider.IRangeInfo)))
		{
			return new CompileResult(eErrorType.Value);
		}
		double resultNumeric = l.ResultNumeric;
		double resultNumeric2 = r.ResultNumeric;
		if (Math.Abs(resultNumeric2 - 0.0) < double.Epsilon)
		{
			return new CompileResult(eErrorType.Div0);
		}
		return ((l.IsNumeric || l.IsNumericString || l.IsDateString || l.Result is ExcelDataProvider.IRangeInfo) && (r.IsNumeric || r.IsNumericString || r.IsDateString || r.Result is ExcelDataProvider.IRangeInfo)) ? new CompileResult(resultNumeric / resultNumeric2, DataType.Decimal) : new CompileResult(eErrorType.Value);
	}));

	public static IOperator Exp => new Operator(Operators.Exponentiation, 4, delegate(CompileResult l, CompileResult r)
	{
		if (l == null && r == null)
		{
			return new CompileResult(eErrorType.Value);
		}
		l = l ?? new CompileResult(0, DataType.Integer);
		r = r ?? new CompileResult(0, DataType.Integer);
		return ((l.IsNumeric || l.IsNumericString || l.IsDateString || l.Result is ExcelDataProvider.IRangeInfo) && (r.IsNumeric || r.IsNumericString || r.IsDateString || r.Result is ExcelDataProvider.IRangeInfo)) ? new CompileResult(Math.Pow(l.ResultNumeric, r.ResultNumeric), DataType.Decimal) : new CompileResult(0.0, DataType.Decimal);
	});

	public static IOperator Concat => new Operator(Operators.Concat, 15, delegate(CompileResult l, CompileResult r)
	{
		l = l ?? new CompileResult(string.Empty, DataType.String);
		r = r ?? new CompileResult(string.Empty, DataType.String);
		string obj = ((l.Result != null) ? l.ResultValue.ToString() : string.Empty);
		string text = ((r.Result != null) ? r.ResultValue.ToString() : string.Empty);
		return new CompileResult(obj + text, DataType.String);
	});

	public static IOperator GreaterThan => _greaterThan ?? (_greaterThan = new Operator(Operators.GreaterThan, 25, (CompileResult l, CompileResult r) => Compare(l, r, (int compRes) => compRes > 0)));

	public static IOperator Eq => _eq ?? (_eq = new Operator(Operators.Equals, 25, (CompileResult l, CompileResult r) => Compare(l, r, (int compRes) => compRes == 0)));

	public static IOperator NotEqualsTo => _notEqualsTo ?? (_notEqualsTo = new Operator(Operators.NotEqualTo, 25, (CompileResult l, CompileResult r) => Compare(l, r, (int compRes) => compRes != 0)));

	public static IOperator GreaterThanOrEqual => _greaterThanOrEqual ?? (_greaterThanOrEqual = new Operator(Operators.GreaterThanOrEqual, 25, (CompileResult l, CompileResult r) => Compare(l, r, (int compRes) => compRes >= 0)));

	public static IOperator LessThan => _lessThan ?? (_lessThan = new Operator(Operators.LessThan, 25, (CompileResult l, CompileResult r) => Compare(l, r, (int compRes) => compRes < 0)));

	public static IOperator LessThanOrEqual => new Operator(Operators.LessThanOrEqual, 25, (CompileResult l, CompileResult r) => Compare(l, r, (int compRes) => compRes <= 0));

	public static IOperator Percent
	{
		get
		{
			if (_percent == null)
			{
				_percent = new Operator(Operators.Percent, 2, delegate(CompileResult l, CompileResult r)
				{
					l = l ?? new CompileResult(0, DataType.Integer);
					r = r ?? new CompileResult(0, DataType.Integer);
					if (l.DataType == DataType.Integer && r.DataType == DataType.Integer)
					{
						return new CompileResult(l.ResultNumeric * r.ResultNumeric, DataType.Integer);
					}
					return ((l.IsNumeric || l.IsNumericString || l.IsDateString || l.Result is ExcelDataProvider.IRangeInfo) && (r.IsNumeric || r.IsNumericString || r.IsDateString || r.Result is ExcelDataProvider.IRangeInfo)) ? new CompileResult(l.ResultNumeric * r.ResultNumeric, DataType.Decimal) : new CompileResult(eErrorType.Value);
				});
			}
			return _percent;
		}
	}

	private Operator()
	{
	}

	private Operator(Operators @operator, int precedence, Func<CompileResult, CompileResult, CompileResult> implementation)
	{
		_implementation = implementation;
		_precedence = precedence;
		_operator = @operator;
	}

	public CompileResult Apply(CompileResult left, CompileResult right)
	{
		if (left.Result is ExcelErrorValue)
		{
			return new CompileResult(left.Result, DataType.ExcelError);
		}
		if (right.Result is ExcelErrorValue)
		{
			return new CompileResult(right.Result, DataType.ExcelError);
		}
		return _implementation(left, right);
	}

	public override string ToString()
	{
		return "Operator: " + _operator;
	}

	private static object GetObjFromOther(CompileResult obj, CompileResult other)
	{
		if (obj.Result == null)
		{
			if (other.DataType == DataType.String)
			{
				return string.Empty;
			}
			return 0.0;
		}
		return obj.ResultValue;
	}

	private static CompileResult Compare(CompileResult l, CompileResult r, Func<int, bool> comparison)
	{
		if (EitherIsError(l, r, out var errorVal))
		{
			return new CompileResult(errorVal);
		}
		object objFromOther = GetObjFromOther(l, r);
		object objFromOther2 = GetObjFromOther(r, l);
		if (ConvertUtil.IsNumeric(objFromOther) && ConvertUtil.IsNumeric(objFromOther2))
		{
			double valueDouble = ConvertUtil.GetValueDouble(objFromOther);
			double valueDouble2 = ConvertUtil.GetValueDouble(objFromOther2);
			if (Math.Abs(valueDouble - valueDouble2) < double.Epsilon)
			{
				return new CompileResult(comparison(0), DataType.Boolean);
			}
			int arg = valueDouble.CompareTo(valueDouble2);
			return new CompileResult(comparison(arg), DataType.Boolean);
		}
		int arg2 = CompareString(objFromOther, objFromOther2);
		return new CompileResult(comparison(arg2), DataType.Boolean);
	}

	private static int CompareString(object l, object r)
	{
		string strA = (l ?? "").ToString();
		string strB = (r ?? "").ToString();
		return string.Compare(strA, strB, StringComparison.OrdinalIgnoreCase);
	}

	private static bool EitherIsError(CompileResult l, CompileResult r, out ExcelErrorValue errorVal)
	{
		if (l.DataType == DataType.ExcelError)
		{
			errorVal = (ExcelErrorValue)l.Result;
			return true;
		}
		if (r.DataType == DataType.ExcelError)
		{
			errorVal = (ExcelErrorValue)r.Result;
			return true;
		}
		errorVal = null;
		return false;
	}
}
