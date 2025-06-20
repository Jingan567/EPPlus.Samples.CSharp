using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime.Workdays;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;

public class NetworkdaysIntl : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		FunctionArgument[] array = (arguments as FunctionArgument[]) ?? arguments.ToArray();
		ValidateArguments(array, 2);
		System.DateTime startDate = System.DateTime.FromOADate(ArgToInt(array, 0));
		System.DateTime endDate = System.DateTime.FromOADate(ArgToInt(array, 1));
		WorkdayCalculator workdayCalculator = new WorkdayCalculator();
		HolidayWeekdaysFactory holidayWeekdaysFactory = new HolidayWeekdaysFactory();
		if (array.Length > 2)
		{
			object value = array[2].Value;
			if (Regex.IsMatch(value.ToString(), "^[01]{7}"))
			{
				workdayCalculator = new WorkdayCalculator(holidayWeekdaysFactory.Create(value.ToString()));
			}
			else
			{
				if (!IsNumeric(value))
				{
					return new CompileResult(eErrorType.Value);
				}
				int code = Convert.ToInt32(value);
				workdayCalculator = new WorkdayCalculator(holidayWeekdaysFactory.Create(code));
			}
		}
		WorkdayCalculatorResult workdayCalculatorResult = workdayCalculator.CalculateNumberOfWorkdays(startDate, endDate);
		if (array.Length > 3)
		{
			workdayCalculatorResult = workdayCalculator.ReduceWorkdaysWithHolidays(workdayCalculatorResult, array[3]);
		}
		return new CompileResult(workdayCalculatorResult.NumberOfWorkdays, DataType.Integer);
	}
}
