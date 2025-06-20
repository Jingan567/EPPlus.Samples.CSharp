using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;

public class Days360 : ExcelFunction
{
	private enum Days360Calctype
	{
		European,
		Us
	}

	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 2);
		double d = ArgToDecimal(arguments, 0);
		double d2 = ArgToDecimal(arguments, 1);
		System.DateTime dateTime = System.DateTime.FromOADate(d);
		System.DateTime dateTime2 = System.DateTime.FromOADate(d2);
		Days360Calctype days360Calctype = Days360Calctype.Us;
		if (arguments.Count() > 2 && ArgToBool(arguments, 2))
		{
			days360Calctype = Days360Calctype.European;
		}
		int year = dateTime.Year;
		int month = dateTime.Month;
		int num = dateTime.Day;
		int year2 = dateTime2.Year;
		int month2 = dateTime2.Month;
		int num2 = dateTime2.Day;
		if (days360Calctype == Days360Calctype.European)
		{
			if (num == 31)
			{
				num = 30;
			}
			if (num2 == 31)
			{
				num2 = 30;
			}
		}
		else
		{
			int num3 = (new GregorianCalendar().IsLeapYear(dateTime.Year) ? 29 : 28);
			if (month == 2 && num == num3 && month2 == 2 && num2 == num3)
			{
				num2 = 30;
			}
			if (month == 2 && num == num3)
			{
				num = 30;
			}
			if (num2 == 31 && (num == 30 || num == 31))
			{
				num2 = 30;
			}
			if (num == 31)
			{
				num = 30;
			}
		}
		int num4 = year2 * 12 * 30 + month2 * 30 + num2 - (year * 12 * 30 + month * 30 + num);
		return CreateResult(num4, DataType.Integer);
	}

	private int GetNumWholeMonths(System.DateTime dt1, System.DateTime dt2)
	{
		System.DateTime dateTime = new System.DateTime(dt1.Year, dt1.Month, 1).AddMonths(1);
		System.DateTime dateTime2 = new System.DateTime(dt2.Year, dt2.Month, 1);
		return (dateTime2.Year - dateTime.Year) * 12 + (dateTime2.Month - dateTime.Month);
	}
}
