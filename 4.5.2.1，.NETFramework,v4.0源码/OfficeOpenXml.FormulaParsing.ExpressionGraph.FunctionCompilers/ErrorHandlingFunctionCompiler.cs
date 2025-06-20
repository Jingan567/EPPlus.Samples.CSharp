using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;

public class ErrorHandlingFunctionCompiler : FunctionCompiler
{
	public ErrorHandlingFunctionCompiler(ExcelFunction function)
		: base(function)
	{
	}

	public override CompileResult Compile(IEnumerable<Expression> children, ParsingContext context)
	{
		List<FunctionArgument> list = new List<FunctionArgument>();
		base.Function.BeforeInvoke(context);
		foreach (Expression child in children)
		{
			try
			{
				BuildFunctionArguments(child.Compile()?.Result, list);
			}
			catch (ExcelErrorValueException ex)
			{
				return ((ErrorHandlingFunction)base.Function).HandleError(ex.ErrorValue.ToString());
			}
			catch
			{
				return ((ErrorHandlingFunction)base.Function).HandleError("#VALUE!");
			}
		}
		return base.Function.Execute(list, context);
	}
}
