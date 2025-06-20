using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime.Workdays;

public class HolidayWeekdays
{
	private readonly List<DayOfWeek> _holidayDays = new List<DayOfWeek>();

	public int NumberOfWorkdaysPerWeek => 7 - _holidayDays.Count;

	public HolidayWeekdays()
		: this(DayOfWeek.Saturday, DayOfWeek.Sunday)
	{
	}

	public HolidayWeekdays(params DayOfWeek[] holidayDays)
	{
		foreach (DayOfWeek item in holidayDays)
		{
			_holidayDays.Add(item);
		}
	}

	public bool IsHolidayWeekday(System.DateTime dateTime)
	{
		return _holidayDays.Contains(dateTime.DayOfWeek);
	}

	public System.DateTime AdjustResultWithHolidays(System.DateTime resultDate, IEnumerable<FunctionArgument> arguments)
	{
		if (arguments.Count() == 2)
		{
			return resultDate;
		}
		if (arguments.ElementAt(2).Value is IEnumerable<FunctionArgument> enumerable)
		{
			foreach (FunctionArgument item in enumerable)
			{
				if (ConvertUtil.IsNumeric(item.Value))
				{
					System.DateTime dateTime = System.DateTime.FromOADate(ConvertUtil.GetValueDouble(item.Value));
					if (!IsHolidayWeekday(dateTime))
					{
						resultDate = resultDate.AddDays(1.0);
					}
				}
			}
		}
		else if (arguments.ElementAt(2).Value is ExcelDataProvider.IRangeInfo rangeInfo)
		{
			foreach (ExcelDataProvider.ICellInfo item2 in rangeInfo)
			{
				if (ConvertUtil.IsNumeric(item2.Value))
				{
					System.DateTime dateTime2 = System.DateTime.FromOADate(ConvertUtil.GetValueDouble(item2.Value));
					if (!IsHolidayWeekday(dateTime2))
					{
						resultDate = resultDate.AddDays(1.0);
					}
				}
			}
		}
		return resultDate;
	}

	public System.DateTime GetNextWorkday(System.DateTime date, WorkdayCalculationDirection direction = WorkdayCalculationDirection.Forward)
	{
		System.DateTime dateTime = date.AddDays((double)direction);
		while (IsHolidayWeekday(dateTime))
		{
			dateTime = dateTime.AddDays((double)direction);
		}
		return dateTime;
	}
}
