namespace OfficeOpenXml.FormulaParsing.ExpressionGraph;

public class BooleanExpression : AtomicExpression
{
	private bool? _precompiledValue;

	public BooleanExpression(string expression)
		: base(expression)
	{
	}

	public BooleanExpression(bool value)
		: base(value ? "true" : "false")
	{
		_precompiledValue = value;
	}

	public override CompileResult Compile()
	{
		return new CompileResult(_precompiledValue ?? bool.Parse(base.ExpressionString), DataType.Boolean);
	}
}
