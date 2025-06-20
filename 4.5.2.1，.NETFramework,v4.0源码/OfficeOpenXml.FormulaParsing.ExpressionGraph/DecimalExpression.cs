using System.Globalization;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph;

public class DecimalExpression : AtomicExpression
{
	private readonly double? _compiledValue;

	private readonly bool _negate;

	public DecimalExpression(string expression)
		: this(expression, negate: false)
	{
	}

	public DecimalExpression(string expression, bool negate)
		: base(expression)
	{
		_negate = negate;
	}

	public DecimalExpression(double compiledValue)
		: base(compiledValue.ToString(CultureInfo.InvariantCulture))
	{
		_compiledValue = compiledValue;
	}

	public override CompileResult Compile()
	{
		double num = (_compiledValue.HasValue ? _compiledValue.Value : double.Parse(base.ExpressionString, CultureInfo.InvariantCulture));
		num = (_negate ? (num * -1.0) : num);
		return new CompileResult(num, DataType.Decimal);
	}
}
