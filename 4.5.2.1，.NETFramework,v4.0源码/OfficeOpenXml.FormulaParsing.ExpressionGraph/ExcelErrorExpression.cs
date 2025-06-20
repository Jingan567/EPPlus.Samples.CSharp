namespace OfficeOpenXml.FormulaParsing.ExpressionGraph;

public class ExcelErrorExpression : Expression
{
	private ExcelErrorValue _error;

	public override bool IsGroupedExpression => false;

	public ExcelErrorExpression(string expression, ExcelErrorValue error)
		: base(expression)
	{
		_error = error;
	}

	public ExcelErrorExpression(ExcelErrorValue error)
		: this(error.ToString(), error)
	{
	}

	public override CompileResult Compile()
	{
		return new CompileResult(_error, DataType.ExcelError);
	}
}
