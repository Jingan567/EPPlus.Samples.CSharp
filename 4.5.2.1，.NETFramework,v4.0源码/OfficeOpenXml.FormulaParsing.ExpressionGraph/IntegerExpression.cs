using System;
using System.Globalization;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph;

public class IntegerExpression : AtomicExpression
{
	private readonly double? _compiledValue;

	private readonly bool _negate;

	public IntegerExpression(string expression)
		: this(expression, negate: false)
	{
	}

	public IntegerExpression(string expression, bool negate)
		: base(expression)
	{
		_negate = negate;
	}

	public IntegerExpression(double val)
		: base(val.ToString(CultureInfo.InvariantCulture))
	{
		_compiledValue = Math.Floor(val);
	}

	public override CompileResult Compile()
	{
		double num = (_compiledValue.HasValue ? _compiledValue.Value : double.Parse(base.ExpressionString, CultureInfo.InvariantCulture));
		num = (_negate ? (num * -1.0) : num);
		return new CompileResult(num, DataType.Integer);
	}
}
