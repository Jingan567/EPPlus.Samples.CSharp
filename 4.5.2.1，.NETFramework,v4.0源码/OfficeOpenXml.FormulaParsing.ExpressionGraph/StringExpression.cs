namespace OfficeOpenXml.FormulaParsing.ExpressionGraph;

public class StringExpression : AtomicExpression
{
	public StringExpression(string expression)
		: base(expression)
	{
	}

	public override CompileResult Compile()
	{
		return new CompileResult(base.ExpressionString, DataType.String);
	}
}
