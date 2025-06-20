using System.Text.RegularExpressions;

namespace OfficeOpenXml.Utils;

public static class AddressUtility
{
	public static string ParseEntireColumnSelections(string address)
	{
		string address2 = address;
		foreach (Match item in Regex.Matches(address, "[A-Z]+:[A-Z]+"))
		{
			AddRowNumbersToEntireColumnRange(ref address2, item.Value);
		}
		return address2;
	}

	private static void AddRowNumbersToEntireColumnRange(ref string address, string range)
	{
		string[] array = $"{range}{1048576}".Split(':');
		address = address.Replace(range, $"{array[0]}1:{array[1]}");
	}
}
