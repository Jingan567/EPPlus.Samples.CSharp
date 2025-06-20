using System;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime.Workdays;

public class WorkdayCalculator
{
	private readonly HolidayWeekdays _holidayWeekdays;

	public WorkdayCalculator()
		: this(new HolidayWeekdays())
	{
	}

	public WorkdayCalculator(HolidayWeekdays holidayWeekdays)
	{
		_holidayWeekdays = holidayWeekdays;
	}

	public WorkdayCalculatorResult CalculateNumberOfWorkdays(System.DateTime startDate, System.DateTime endDate)
	{
		WorkdayCalculationDirection workdayCalculationDirection = ((startDate < endDate) ? WorkdayCalculationDirection.Forward : WorkdayCalculationDirection.Backward);
		System.DateTime date;
		System.DateTime date2;
		if (workdayCalculationDirection == WorkdayCalculationDirection.Forward)
		{
			date = startDate.Date;
			date2 = endDate.Date;
		}
		else
		{
			date = endDate.Date;
			date2 = startDate.Date;
		}
		int num = (int)date2.Subtract(date).TotalDays / 7;
		int num2 = num * _holidayWeekdays.NumberOfWorkdaysPerWeek;
		if (!_holidayWeekdays.IsHolidayWeekday(date))
		{
			num2++;
		}
		System.DateTime dateTime = date.AddDays(num * 7);
		while (dateTime < date2)
		{
			dateTime = dateTime.AddDays(1.0);
			if (!_holidayWeekdays.IsHolidayWeekday(dateTime))
			{
				num2++;
			}
		}
		return new WorkdayCalculatorResult(num2, startDate, endDate, workdayCalculationDirection);
	}

	public WorkdayCalculatorResult CalculateWorkday(System.DateTime startDate, int nWorkDays)
	{
		WorkdayCalculationDirection workdayCalculationDirection = ((nWorkDays > 0) ? WorkdayCalculationDirection.Forward : WorkdayCalculationDirection.Backward);
		int num = (int)workdayCalculationDirection;
		nWorkDays *= num;
		int num2 = 0;
		System.DateTime dateTime = startDate;
		int num3 = nWorkDays / _holidayWeekdays.NumberOfWorkdaysPerWeek;
		dateTime = dateTime.AddDays(num3 * 7 * num);
		num2 += num3 * _holidayWeekdays.NumberOfWorkdaysPerWeek;
		while (num2 < nWorkDays)
		{
			dateTime = dateTime.AddDays(num);
			if (!_holidayWeekdays.IsHolidayWeekday(dateTime))
			{
				num2++;
			}
		}
		return new WorkdayCalculatorResult(num2, startDate, dateTime, workdayCalculationDirection);
	}

	public WorkdayCalculatorResult ReduceWorkdaysWithHolidays(WorkdayCalculatorResult calculatedResult, FunctionArgument holidayArgument)
	{
		System.DateTime startDate = calculatedResult.StartDate;
		System.DateTime endDate = calculatedResult.EndDate;
		AdditionalHolidayDays additionalHolidayDays = new AdditionalHolidayDays(holidayArgument);
		System.DateTime calcStartDate;
		System.DateTime calcEndDate;
		if (startDate < endDate)
		{
			calcStartDate = startDate;
			calcEndDate = endDate;
		}
		else
		{
			calcStartDate = endDate;
			calcEndDate = startDate;
		}
		int num = additionalHolidayDays.AdditionalDates.Count((System.DateTime x) => x >= calcStartDate && x <= calcEndDate && !_holidayWeekdays.IsHolidayWeekday(x));
		return new WorkdayCalculatorResult(calculatedResult.NumberOfWorkdays - num, startDate, endDate, calculatedResult.Direction);
	}

	public WorkdayCalculatorResult AdjustResultWithHolidays(WorkdayCalculatorResult calculatedResult, FunctionArgument holidayArgument)
	{
		System.DateTime startDate = calculatedResult.StartDate;
		System.DateTime dateTime = calculatedResult.EndDate;
		WorkdayCalculationDirection direction = calculatedResult.Direction;
		int num = calculatedResult.NumberOfWorkdays;
		AdditionalHolidayDays additionalHolidayDays = new AdditionalHolidayDays(holidayArgument);
		foreach (System.DateTime additionalDate in additionalHolidayDays.AdditionalDates)
		{
			if ((direction != WorkdayCalculationDirection.Forward || (!(additionalDate < startDate) && !(additionalDate > dateTime))) && (direction != WorkdayCalculationDirection.Backward || (!(additionalDate > startDate) && !(additionalDate < dateTime))) && !_holidayWeekdays.IsHolidayWeekday(additionalDate))
			{
				System.DateTime nextWorkday = _holidayWeekdays.GetNextWorkday(dateTime, direction);
				while (additionalHolidayDays.AdditionalDates.Contains(nextWorkday))
				{
					nextWorkday = _holidayWeekdays.GetNextWorkday(nextWorkday, direction);
				}
				num++;
				dateTime = nextWorkday;
			}
		}
		return new WorkdayCalculatorResult(num, calculatedResult.StartDate, dateTime, direction);
	}
}
