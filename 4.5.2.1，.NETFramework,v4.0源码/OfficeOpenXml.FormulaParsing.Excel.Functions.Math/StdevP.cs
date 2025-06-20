using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class StdevP : HiddenValuesHandlingFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		IEnumerable<ExcelDoubleCellValue> source = ArgsToDoubleEnumerable(base.IgnoreHiddenValues, ignoreErrors: false, arguments, context);
		return CreateResult(StandardDeviation(source.Select((Func<ExcelDoubleCellValue, double>)((ExcelDoubleCellValue x) => x))), DataType.Decimal);
	}

	private static double StandardDeviation(IEnumerable<double> values)
	{
		double avg = values.Average();
		return System.Math.Sqrt(values.Average((double v) => System.Math.Pow(v - avg, 2.0)));
	}
}
