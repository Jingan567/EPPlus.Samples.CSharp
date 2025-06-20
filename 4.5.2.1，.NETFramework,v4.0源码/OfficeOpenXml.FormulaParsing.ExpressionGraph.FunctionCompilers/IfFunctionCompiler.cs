using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;

public class IfFunctionCompiler : FunctionCompiler
{
	public IfFunctionCompiler(ExcelFunction function)
		: base(function)
	{
		OfficeOpenXml.FormulaParsing.Utilities.Require.That(function).Named("function").IsNotNull();
		if (!(function is If))
		{
			throw new ArgumentException("function must be of type If");
		}
	}

	public override CompileResult Compile(IEnumerable<Expression> children, ParsingContext context)
	{
		if (children.Count() < 2)
		{
			throw new ExcelErrorValueException(eErrorType.Value);
		}
		List<FunctionArgument> list = new List<FunctionArgument>();
		base.Function.BeforeInvoke(context);
		object obj = children.ElementAt(0).Compile().Result;
		if (obj is ExcelDataProvider.INameInfo)
		{
			obj = ((ExcelDataProvider.INameInfo)obj).Value;
		}
		if (obj is ExcelDataProvider.IRangeInfo)
		{
			ExcelDataProvider.IRangeInfo obj2 = (ExcelDataProvider.IRangeInfo)obj;
			if (obj2.GetNCells() > 1)
			{
				throw new ArgumentException("Logical can't be more than one cell");
			}
			obj = obj2.GetOffset(0, 0);
		}
		bool result;
		if (obj is bool)
		{
			result = (bool)obj;
		}
		else if (!ConvertUtil.TryParseBooleanString(obj, out result))
		{
			if (!ConvertUtil.IsNumeric(obj))
			{
				throw new ArgumentException("Invalid logical test");
			}
			result = ConvertUtil.GetValueDouble(obj) != 0.0;
		}
		list.Add(new FunctionArgument(result));
		if (result)
		{
			object result2 = children.ElementAt(1).Compile().Result;
			list.Add(new FunctionArgument(result2));
			list.Add(new FunctionArgument(null));
		}
		else
		{
			Expression expression = children.ElementAtOrDefault(2);
			object val = ((expression != null) ? expression.Compile().Result : ((object)false));
			list.Add(new FunctionArgument(null));
			list.Add(new FunctionArgument(val));
		}
		return base.Function.Execute(list, context);
	}
}
