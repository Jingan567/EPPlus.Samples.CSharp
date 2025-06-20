using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class Median : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		IEnumerable<ExcelDoubleCellValue> source = ArgsToDoubleEnumerable(arguments, context);
		ExcelDoubleCellValue[] arr = source.ToArray();
		Array.Sort(arr);
		ThrowExcelErrorValueExceptionIf(() => arr.Length == 0, eErrorType.Num);
		double num;
		if (arr.Length % 2 == 1)
		{
			num = arr[arr.Length / 2];
		}
		else
		{
			int num2 = arr.Length / 2 - 1;
			num = ((double)arr[num2] + (double)arr[num2 + 1]) / 2.0;
		}
		return CreateResult(num, DataType.Decimal);
	}
}
