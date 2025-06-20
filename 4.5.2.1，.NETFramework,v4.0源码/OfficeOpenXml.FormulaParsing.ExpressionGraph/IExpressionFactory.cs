using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph;

public interface IExpressionFactory
{
	Expression Create(Token token);
}
