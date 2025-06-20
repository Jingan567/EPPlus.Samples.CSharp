using System;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph;

public class DateExpression : AtomicExpression
{
	public DateExpression(string expression)
		: base(expression)
	{
	}

	public override CompileResult Compile()
	{
		return new CompileResult(DateTime.FromOADate(double.Parse(base.ExpressionString)), DataType.Date);
	}
}
