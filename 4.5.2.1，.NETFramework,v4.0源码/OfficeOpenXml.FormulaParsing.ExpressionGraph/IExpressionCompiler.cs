using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph;

public interface IExpressionCompiler
{
	CompileResult Compile(IEnumerable<Expression> expressions);
}
