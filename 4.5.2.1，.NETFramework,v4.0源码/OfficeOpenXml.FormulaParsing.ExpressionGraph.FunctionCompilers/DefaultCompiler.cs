using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.Excel;
using OfficeOpenXml.FormulaParsing.Excel.Functions;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;

public class DefaultCompiler : FunctionCompiler
{
	public DefaultCompiler(ExcelFunction function)
		: base(function)
	{
	}

	public override CompileResult Compile(IEnumerable<Expression> children, ParsingContext context)
	{
		List<FunctionArgument> list = new List<FunctionArgument>();
		base.Function.BeforeInvoke(context);
		foreach (Expression child in children)
		{
			CompileResult compileResult = child.Compile();
			if (compileResult.IsResultOfSubtotal)
			{
				FunctionArgument functionArgument = new FunctionArgument(compileResult.Result);
				functionArgument.SetExcelStateFlag(ExcelCellState.IsResultOfSubtotal);
				list.Add(functionArgument);
			}
			else
			{
				BuildFunctionArguments(compileResult.Result, list);
			}
		}
		return base.Function.Execute(list, context);
	}
}
