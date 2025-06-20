namespace OfficeOpenXml.Utils;

public static class Require
{
	public static IArgument<T> Argument<T>(T argument)
	{
		return new Argument<T>(argument);
	}
}
