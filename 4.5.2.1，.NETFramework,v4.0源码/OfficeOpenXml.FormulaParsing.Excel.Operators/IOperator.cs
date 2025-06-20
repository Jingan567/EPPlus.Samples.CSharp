using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Operators;

public interface IOperator
{
	Operators Operator { get; }

	int Precedence { get; }

	CompileResult Apply(CompileResult left, CompileResult right);
}
