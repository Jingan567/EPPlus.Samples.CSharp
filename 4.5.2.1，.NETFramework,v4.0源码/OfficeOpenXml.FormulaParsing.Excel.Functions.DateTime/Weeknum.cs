using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;

public class Weeknum : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1, eErrorType.Value);
		System.DateTime time = System.DateTime.FromOADate(ArgToDecimal(arguments, 0));
		DayOfWeek firstDayOfWeek = DayOfWeek.Sunday;
		if (arguments.Count() > 1)
		{
			switch (ArgToInt(arguments, 1))
			{
			case 1:
				firstDayOfWeek = DayOfWeek.Sunday;
				break;
			case 2:
			case 11:
				firstDayOfWeek = DayOfWeek.Monday;
				break;
			case 12:
				firstDayOfWeek = DayOfWeek.Tuesday;
				break;
			case 13:
				firstDayOfWeek = DayOfWeek.Wednesday;
				break;
			case 14:
				firstDayOfWeek = DayOfWeek.Thursday;
				break;
			case 15:
				firstDayOfWeek = DayOfWeek.Friday;
				break;
			case 16:
				firstDayOfWeek = DayOfWeek.Saturday;
				break;
			default:
				ThrowExcelErrorValueException(eErrorType.Num);
				break;
			}
		}
		if (DateTimeFormatInfo.CurrentInfo == null)
		{
			throw new InvalidOperationException("Could not execute Weeknum function because DateTimeFormatInfo.CurrentInfo was null");
		}
		int weekOfYear = DateTimeFormatInfo.CurrentInfo.Calendar.GetWeekOfYear(time, CalendarWeekRule.FirstDay, firstDayOfWeek);
		return CreateResult(weekOfYear, DataType.Integer);
	}
}
