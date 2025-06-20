using System.Text.RegularExpressions;

namespace OfficeOpenXml.Utils;

public static class SqRefUtility
{
	public static string ToSqRefAddress(string address)
	{
		Require.Argument(address).IsNotNullOrEmpty(address);
		address = address.Replace(",", " ");
		address = new Regex("[ ]+").Replace(address, " ");
		return address;
	}

	public static string FromSqRefAddress(string address)
	{
		Require.Argument(address).IsNotNullOrEmpty(address);
		address = address.Replace(" ", ",");
		return address;
	}
}
