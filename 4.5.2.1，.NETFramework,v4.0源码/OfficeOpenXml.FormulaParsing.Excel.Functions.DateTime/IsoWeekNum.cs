using System;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;

public class IsoWeekNum : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		System.DateTime fromDate = System.DateTime.FromOADate(ArgToInt(arguments, 0));
		return CreateResult(WeekNumber(fromDate), DataType.Integer);
	}

	private int WeekNumber(System.DateTime fromDate)
	{
		System.DateTime value = fromDate.AddDays(-fromDate.Day + 1).AddMonths(-fromDate.Month + 1);
		System.DateTime dateTime = value.AddYears(1).AddDays(-1.0);
		int[] array = new int[7] { 6, 7, 8, 9, 10, 4, 5 };
		int num = (fromDate.Subtract(value).Days + array[(int)value.DayOfWeek]) / 7;
		switch (num)
		{
		case 0:
			return WeekNumber(value.AddDays(-1.0));
		case 53:
			if (dateTime.DayOfWeek < DayOfWeek.Thursday)
			{
				return 1;
			}
			return num;
		default:
			return num;
		}
	}
}
