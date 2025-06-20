namespace OfficeOpenXml.FormulaParsing.ExcelUtilities;

public static class ExcelAddressUtil
{
	private static char[] SheetNameInvalidChars = new char[5] { '?', ':', '*', '/', '\\' };

	public static bool IsValidAddress(string token)
	{
		int num;
		if (token[0] == '\'')
		{
			num = token.LastIndexOf('\'');
			if (num <= 0 || num >= token.Length - 1 || token[num + 1] != '!')
			{
				return false;
			}
			if (token.IndexOfAny(SheetNameInvalidChars, 1, num - 1) > 0)
			{
				return false;
			}
			token = token.Substring(num + 2);
		}
		else if ((num = token.IndexOf('!')) > 1)
		{
			if (token.IndexOfAny(SheetNameInvalidChars, 0, token.IndexOf('!')) > 0)
			{
				return false;
			}
			token = token.Substring(token.IndexOf('!') + 1);
		}
		return ExcelCellBase.IsValidAddress(token);
	}
}
