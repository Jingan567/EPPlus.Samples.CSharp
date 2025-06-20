using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime.Workdays;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;

public class Networkdays : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		FunctionArgument[] array = (arguments as FunctionArgument[]) ?? arguments.ToArray();
		ValidateArguments(array, 2);
		System.DateTime startDate = System.DateTime.FromOADate(ArgToInt(array, 0));
		System.DateTime endDate = System.DateTime.FromOADate(ArgToInt(array, 1));
		WorkdayCalculator workdayCalculator = new WorkdayCalculator();
		WorkdayCalculatorResult workdayCalculatorResult = workdayCalculator.CalculateNumberOfWorkdays(startDate, endDate);
		if (array.Length > 2)
		{
			workdayCalculatorResult = workdayCalculator.ReduceWorkdaysWithHolidays(workdayCalculatorResult, array[2]);
		}
		return new CompileResult(workdayCalculatorResult.NumberOfWorkdays, DataType.Integer);
	}
}
