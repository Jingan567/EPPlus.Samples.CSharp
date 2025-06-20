using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph;

public interface IExpressionGraphBuilder
{
	ExpressionGraph Build(IEnumerable<Token> tokens);
}
