using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;

public abstract class FunctionCompiler
{
	protected ExcelFunction Function { get; private set; }

	public FunctionCompiler(ExcelFunction function)
	{
		Require.That(function).Named("function").IsNotNull();
		Function = function;
	}

	protected void BuildFunctionArguments(object result, DataType dataType, List<FunctionArgument> args)
	{
		if (result is IEnumerable<object> && !(result is ExcelDataProvider.IRangeInfo))
		{
			List<FunctionArgument> list = new List<FunctionArgument>();
			foreach (object item in result as IEnumerable<object>)
			{
				BuildFunctionArguments(item, dataType, list);
			}
			args.Add(new FunctionArgument(list));
		}
		else
		{
			args.Add(new FunctionArgument(result, dataType));
		}
	}

	protected void BuildFunctionArguments(object result, List<FunctionArgument> args)
	{
		BuildFunctionArguments(result, DataType.Unknown, args);
	}

	public abstract CompileResult Compile(IEnumerable<Expression> children, ParsingContext context);
}
