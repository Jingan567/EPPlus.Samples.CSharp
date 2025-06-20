namespace OfficeOpenXml.FormulaParsing.ExpressionGraph;

public interface IExpressionConverter
{
	StringExpression ToStringExpression(Expression expression);

	Expression FromCompileResult(CompileResult compileResult);
}
