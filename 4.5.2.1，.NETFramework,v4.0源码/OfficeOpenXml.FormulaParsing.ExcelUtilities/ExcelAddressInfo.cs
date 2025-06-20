using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities;

public class ExcelAddressInfo
{
	public string Worksheet { get; private set; }

	public bool WorksheetIsSpecified => !string.IsNullOrEmpty(Worksheet);

	public bool IsMultipleCells => !string.IsNullOrEmpty(EndCell);

	public string StartCell { get; private set; }

	public string EndCell { get; private set; }

	public string AddressOnSheet { get; private set; }

	private ExcelAddressInfo(string address)
	{
		string text = address;
		Worksheet = string.Empty;
		if (address.Contains("!"))
		{
			string[] array = address.Split('!');
			Worksheet = array[0];
			text = array[1];
		}
		if (text.Contains(":"))
		{
			string[] array2 = text.Split(':');
			StartCell = array2[0];
			EndCell = array2[1];
		}
		else
		{
			StartCell = text;
		}
		AddressOnSheet = text;
	}

	public static ExcelAddressInfo Parse(string address)
	{
		Require.That(address).Named("address").IsNotNullOrEmpty();
		return new ExcelAddressInfo(address);
	}
}
