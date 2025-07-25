using System;
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime.Workdays;

public class HolidayWeekdaysFactory
{
	private readonly DayOfWeek[] _dayOfWeekArray = new DayOfWeek[7]
	{
		DayOfWeek.Monday,
		DayOfWeek.Tuesday,
		DayOfWeek.Wednesday,
		DayOfWeek.Thursday,
		DayOfWeek.Friday,
		DayOfWeek.Saturday,
		DayOfWeek.Sunday
	};

	public HolidayWeekdays Create(string weekdays)
	{
		if (string.IsNullOrEmpty(weekdays) || weekdays.Length != 7)
		{
			throw new ArgumentException("Illegal weekday string", "Weekday");
		}
		List<DayOfWeek> list = new List<DayOfWeek>();
		char[] array = weekdays.ToCharArray();
		for (int i = 0; i < array.Length; i++)
		{
			if (array[i] == '1')
			{
				list.Add(_dayOfWeekArray[i]);
			}
		}
		return new HolidayWeekdays(list.ToArray());
	}

	public HolidayWeekdays Create(int code)
	{
		return code switch
		{
			1 => new HolidayWeekdays(DayOfWeek.Saturday, DayOfWeek.Sunday), 
			2 => new HolidayWeekdays(DayOfWeek.Sunday, DayOfWeek.Monday), 
			3 => new HolidayWeekdays(DayOfWeek.Monday, DayOfWeek.Tuesday), 
			4 => new HolidayWeekdays(DayOfWeek.Tuesday, DayOfWeek.Wednesday), 
			5 => new HolidayWeekdays(DayOfWeek.Wednesday, DayOfWeek.Thursday), 
			6 => new HolidayWeekdays(DayOfWeek.Thursday, DayOfWeek.Friday), 
			7 => new HolidayWeekdays(DayOfWeek.Friday, DayOfWeek.Saturday), 
			11 => new HolidayWeekdays(default(DayOfWeek)), 
			12 => new HolidayWeekdays(DayOfWeek.Monday), 
			13 => new HolidayWeekdays(DayOfWeek.Tuesday), 
			14 => new HolidayWeekdays(DayOfWeek.Wednesday), 
			15 => new HolidayWeekdays(DayOfWeek.Thursday), 
			16 => new HolidayWeekdays(DayOfWeek.Friday), 
			17 => new HolidayWeekdays(DayOfWeek.Saturday), 
			_ => throw new ArgumentException("Invalid code supplied to HolidayWeekdaysFactory: " + code), 
		};
	}
}
