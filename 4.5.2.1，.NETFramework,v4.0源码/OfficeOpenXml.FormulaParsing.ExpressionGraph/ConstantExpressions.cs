namespace OfficeOpenXml.FormulaParsing.ExpressionGraph;

public static class ConstantExpressions
{
	public static Expression Percent => new ConstantExpression("Percent", () => new CompileResult(0.01, DataType.Decimal));
}
