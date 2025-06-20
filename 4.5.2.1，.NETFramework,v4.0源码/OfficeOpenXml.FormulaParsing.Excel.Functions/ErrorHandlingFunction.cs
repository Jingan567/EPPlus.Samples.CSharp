using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions;

public abstract class ErrorHandlingFunction : ExcelFunction
{
	public override bool IsErrorHandlingFunction => true;

	public abstract CompileResult HandleError(string errorCode);
}
