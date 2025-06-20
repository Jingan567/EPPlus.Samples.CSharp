using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class CountBlank : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		FunctionArgument functionArgument = arguments.First();
		if (!functionArgument.IsExcelRange)
		{
			throw new InvalidOperationException("CountBlank only support ranges as arguments");
		}
		int num = functionArgument.ValueAsRangeInfo.GetNCells();
		foreach (ExcelDataProvider.ICellInfo item in functionArgument.ValueAsRangeInfo)
		{
			if (item.Value != null && !(item.Value.ToString() == string.Empty))
			{
				num--;
			}
		}
		return CreateResult(num, DataType.Integer);
	}
}
