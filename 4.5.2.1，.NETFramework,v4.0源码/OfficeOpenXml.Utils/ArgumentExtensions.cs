using System;

namespace OfficeOpenXml.Utils;

public static class ArgumentExtensions
{
	public static void IsNotNull<T>(this IArgument<T> argument, string argumentName) where T : class
	{
		argumentName = (string.IsNullOrEmpty(argumentName) ? "value" : argumentName);
		if (argument.Value == null)
		{
			throw new ArgumentNullException(argumentName);
		}
	}

	public static void IsNotNullOrEmpty(this IArgument<string> argument, string argumentName)
	{
		if (string.IsNullOrEmpty(argument.Value))
		{
			throw new ArgumentNullException(argumentName);
		}
	}

	public static void IsInRange<T>(this IArgument<T> argument, T min, T max, string argumentName) where T : IComparable
	{
		if (argument.Value.CompareTo(min) < 0 || argument.Value.CompareTo(max) > 0)
		{
			throw new ArgumentOutOfRangeException(argumentName);
		}
	}
}
