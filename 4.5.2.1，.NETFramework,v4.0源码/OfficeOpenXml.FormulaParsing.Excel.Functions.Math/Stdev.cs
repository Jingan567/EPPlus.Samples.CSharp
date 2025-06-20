using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class Stdev : HiddenValuesHandlingFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		IEnumerable<double> values = ArgsToDoubleEnumerable(arguments, context, ignoreErrors: false).Select((Func<ExcelDoubleCellValue, double>)((ExcelDoubleCellValue x) => x));
		return CreateResult(StandardDeviation(values), DataType.Decimal);
	}

	private double StandardDeviation(IEnumerable<double> values)
	{
		double result = 0.0;
		if (values.Any())
		{
			if (values.Count() == 1)
			{
				throw new ExcelErrorValueException(eErrorType.Div0);
			}
			double avg = values.Average();
			double left = values.Sum((double d) => System.Math.Pow(d - avg, 2.0));
			result = System.Math.Sqrt(Divide(left, values.Count() - 1));
		}
		return result;
	}
}
