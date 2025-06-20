using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;

public class IfErrorFunctionCompiler : FunctionCompiler
{
	public IfErrorFunctionCompiler(ExcelFunction function)
		: base(function)
	{
		Require.That(function).Named("function").IsNotNull();
	}

	public override CompileResult Compile(IEnumerable<Expression> children, ParsingContext context)
	{
		if (children.Count() != 2)
		{
			throw new ExcelErrorValueException(eErrorType.Value);
		}
		List<FunctionArgument> list = new List<FunctionArgument>();
		base.Function.BeforeInvoke(context);
		Expression expression = children.First();
		Expression expression2 = children.ElementAt(1);
		try
		{
			CompileResult compileResult = expression.Compile();
			if (compileResult.DataType == DataType.ExcelError)
			{
				list.Add(new FunctionArgument(expression2.Compile().Result));
			}
			else
			{
				list.Add(new FunctionArgument(compileResult.Result));
			}
		}
		catch (ExcelErrorValueException)
		{
			list.Add(new FunctionArgument(expression2.Compile().Result));
		}
		return base.Function.Execute(list, context);
	}
}
