using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime.Workdays;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;

public class Workday : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		FunctionArgument[] array = (arguments as FunctionArgument[]) ?? arguments.ToArray();
		ValidateArguments(array, 2);
		System.DateTime startDate = System.DateTime.FromOADate(ArgToInt(array, 0));
		int nWorkDays = ArgToInt(array, 1);
		_ = System.DateTime.MinValue;
		WorkdayCalculator workdayCalculator = new WorkdayCalculator();
		WorkdayCalculatorResult workdayCalculatorResult = workdayCalculator.CalculateWorkday(startDate, nWorkDays);
		if (array.Length > 2)
		{
			workdayCalculatorResult = workdayCalculator.AdjustResultWithHolidays(workdayCalculatorResult, array[2]);
		}
		return CreateResult(workdayCalculatorResult.EndDate.ToOADate(), DataType.Date);
	}
}
