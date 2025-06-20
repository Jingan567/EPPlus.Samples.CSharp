namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.CompileStrategy;

public interface ICompileStrategyFactory
{
	CompileStrategy Create(Expression expression);
}
