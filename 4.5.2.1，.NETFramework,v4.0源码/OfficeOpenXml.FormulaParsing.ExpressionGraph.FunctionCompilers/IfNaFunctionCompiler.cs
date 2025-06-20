using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;

public class IfNaFunctionCompiler : FunctionCompiler
{
	public IfNaFunctionCompiler(ExcelFunction function)
		: base(function)
	{
	}

	public override CompileResult Compile(IEnumerable<Expression> children, ParsingContext context)
	{
		if (children.Count() != 2)
		{
			return new CompileResult(eErrorType.Value);
		}
		List<FunctionArgument> list = new List<FunctionArgument>();
		base.Function.BeforeInvoke(context);
		Expression expression = children.First();
		Expression expression2 = children.ElementAt(1);
		try
		{
			CompileResult compileResult = expression.Compile();
			if (compileResult.DataType == DataType.ExcelError && object.Equals(compileResult.Result, ExcelErrorValue.Create(eErrorType.NA)))
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
