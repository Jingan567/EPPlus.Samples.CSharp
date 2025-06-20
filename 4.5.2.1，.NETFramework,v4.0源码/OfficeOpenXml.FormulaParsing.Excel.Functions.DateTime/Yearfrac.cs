using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;

public class Yearfrac : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		FunctionArgument[] array = (arguments as FunctionArgument[]) ?? arguments.ToArray();
		ValidateArguments(array, 2);
		double num = ArgToDecimal(array, 0);
		double num2 = ArgToDecimal(array, 1);
		if (num > num2)
		{
			double num3 = num;
			num = num2;
			num2 = num3;
			FunctionArgument functionArgument = array[1];
			array[1] = array[0];
			array[0] = functionArgument;
		}
		System.DateTime dateTime = System.DateTime.FromOADate(num);
		System.DateTime dateTime2 = System.DateTime.FromOADate(num2);
		int basis = 0;
		if (array.Count() > 2)
		{
			basis = ArgToInt(array, 2);
			ThrowExcelErrorValueExceptionIf(() => basis < 0 || basis > 4, eErrorType.Num);
		}
		ExcelFunction function = context.Configuration.FunctionRepository.GetFunction("days360");
		GregorianCalendar gregorianCalendar = new GregorianCalendar();
		switch (basis)
		{
		case 0:
		{
			double num4 = System.Math.Abs(function.Execute(array, context).ResultNumeric);
			if (dateTime.Month == 2 && dateTime2.Day == 31)
			{
				int num5 = (gregorianCalendar.IsLeapYear(dateTime.Year) ? 29 : 28);
				if (dateTime.Day == num5)
				{
					num4 += 1.0;
				}
			}
			return CreateResult(num4 / 360.0, DataType.Decimal);
		}
		case 1:
			return CreateResult(System.Math.Abs((dateTime2 - dateTime).TotalDays / CalculateAcutalYear(dateTime, dateTime2)), DataType.Decimal);
		case 2:
			return CreateResult(System.Math.Abs((dateTime2 - dateTime).TotalDays / 360.0), DataType.Decimal);
		case 3:
			return CreateResult(System.Math.Abs((dateTime2 - dateTime).TotalDays / 365.0), DataType.Decimal);
		case 4:
		{
			List<FunctionArgument> list = array.ToList();
			list.Add(new FunctionArgument(true));
			return CreateResult(new double?(System.Math.Abs(function.Execute(list, context).ResultNumeric / 360.0)).Value, DataType.Decimal);
		}
		default:
			return null;
		}
	}

	private double CalculateAcutalYear(System.DateTime dt1, System.DateTime dt2)
	{
		GregorianCalendar gregorianCalendar = new GregorianCalendar();
		double num = 0.0;
		int num2 = dt2.Year - dt1.Year + 1;
		for (int i = dt1.Year; i <= dt2.Year; i++)
		{
			num += (double)(gregorianCalendar.IsLeapYear(i) ? 366 : 365);
		}
		if (new System.DateTime(dt1.Year + 1, dt1.Month, dt1.Day) >= dt2)
		{
			num2 = 1;
			num = 365.0;
			if (gregorianCalendar.IsLeapYear(dt1.Year) && dt1.Month <= 2)
			{
				num = 366.0;
			}
			else if (gregorianCalendar.IsLeapYear(dt2.Year) && dt2.Month > 2)
			{
				num = 366.0;
			}
			else if (dt2.Month == 2 && dt2.Day == 29)
			{
				num = 366.0;
			}
		}
		return num / (double)num2;
	}
}
