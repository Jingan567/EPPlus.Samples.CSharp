using System;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;

public class TimeStringParser
{
	private const string RegEx24 = "^[0-9]{1,2}(\\:[0-9]{1,2}){0,2}$";

	private const string RegEx12 = "^[0-9]{1,2}(\\:[0-9]{1,2}){0,2}( PM| AM)$";

	private double GetSerialNumber(int hour, int minute, int second)
	{
		double num = 86400.0;
		return ((double)hour * 60.0 * 60.0 + (double)minute * 60.0 + (double)second) / num;
	}

	private void ValidateValues(int hour, int minute, int second)
	{
		if (second < 0 || second > 59)
		{
			throw new FormatException("Illegal value for second: " + second);
		}
		if (minute < 0 || minute > 59)
		{
			throw new FormatException("Illegal value for minute: " + minute);
		}
	}

	public virtual double Parse(string input)
	{
		return InternalParse(input);
	}

	public virtual bool CanParse(string input)
	{
		System.DateTime result;
		if (!Regex.IsMatch(input, "^[0-9]{1,2}(\\:[0-9]{1,2}){0,2}$") && !Regex.IsMatch(input, "^[0-9]{1,2}(\\:[0-9]{1,2}){0,2}( PM| AM)$"))
		{
			return System.DateTime.TryParse(input, out result);
		}
		return true;
	}

	private double InternalParse(string input)
	{
		if (Regex.IsMatch(input, "^[0-9]{1,2}(\\:[0-9]{1,2}){0,2}$"))
		{
			return Parse24HourTimeString(input);
		}
		if (Regex.IsMatch(input, "^[0-9]{1,2}(\\:[0-9]{1,2}){0,2}( PM| AM)$"))
		{
			return Parse12HourTimeString(input);
		}
		if (System.DateTime.TryParse(input, out var result))
		{
			return GetSerialNumber(result.Hour, result.Minute, result.Second);
		}
		return -1.0;
	}

	private double Parse12HourTimeString(string input)
	{
		_ = string.Empty;
		string text = input.Substring(input.Length - 2, 2);
		GetValuesFromString(input, out var hour, out var minute, out var second);
		if (text == "PM")
		{
			hour += 12;
		}
		ValidateValues(hour, minute, second);
		return GetSerialNumber(hour, minute, second);
	}

	private double Parse24HourTimeString(string input)
	{
		GetValuesFromString(input, out var hour, out var minute, out var second);
		ValidateValues(hour, minute, second);
		return GetSerialNumber(hour, minute, second);
	}

	private static void GetValuesFromString(string input, out int hour, out int minute, out int second)
	{
		hour = 0;
		minute = 0;
		second = 0;
		string[] array = input.Split(':');
		hour = int.Parse(array[0]);
		if (array.Length > 1)
		{
			minute = int.Parse(array[1]);
		}
		if (array.Length > 2)
		{
			string input2 = array[2];
			input2 = Regex.Replace(input2, "[^0-9]+$", string.Empty);
			second = int.Parse(input2);
		}
	}
}
