using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;

public abstract class TimeBaseFunction : ExcelFunction
{
	protected TimeStringParser TimeStringParser { get; private set; }

	protected double SerialNumber { get; private set; }

	protected double SecondsInADay => 86400.0;

	public TimeBaseFunction()
	{
		TimeStringParser = new TimeStringParser();
	}

	public void ValidateAndInitSerialNumber(IEnumerable<FunctionArgument> arguments)
	{
		ValidateArguments(arguments, 1);
		SerialNumber = ArgToDecimal(arguments, 0);
	}

	protected double GetTimeSerialNumber(double seconds)
	{
		return seconds / SecondsInADay;
	}

	protected double GetSeconds(double serialNumber)
	{
		return serialNumber * SecondsInADay;
	}

	protected double GetHour(double serialNumber)
	{
		return (int)GetSeconds(serialNumber) / 3600;
	}

	protected double GetMinute(double serialNumber)
	{
		double num = GetSeconds(serialNumber) - GetHour(serialNumber) * 60.0 * 60.0;
		return (num - num % 60.0) / 60.0;
	}

	protected double GetSecond(double serialNumber)
	{
		return GetSeconds(serialNumber) % 60.0;
	}
}
