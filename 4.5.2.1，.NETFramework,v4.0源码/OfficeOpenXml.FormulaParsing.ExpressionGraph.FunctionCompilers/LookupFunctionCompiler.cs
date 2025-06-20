using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.Excel.Functions;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;

public class LookupFunctionCompiler : FunctionCompiler
{
	public LookupFunctionCompiler(ExcelFunction function)
		: base(function)
	{
	}

	public override CompileResult Compile(IEnumerable<Expression> children, ParsingContext context)
	{
		List<FunctionArgument> list = new List<FunctionArgument>();
		base.Function.BeforeInvoke(context);
		bool flag = true;
		foreach (Expression child in children)
		{
			if (!flag || base.Function.SkipArgumentEvaluation)
			{
				child.ParentIsLookupFunction = base.Function.IsLookupFuction;
			}
			else
			{
				flag = false;
			}
			CompileResult compileResult = child.Compile();
			if (compileResult != null)
			{
				BuildFunctionArguments(compileResult.Result, compileResult.DataType, list);
			}
			else
			{
				BuildFunctionArguments(null, DataType.Unknown, list);
			}
		}
		return base.Function.Execute(list, context);
	}
}
