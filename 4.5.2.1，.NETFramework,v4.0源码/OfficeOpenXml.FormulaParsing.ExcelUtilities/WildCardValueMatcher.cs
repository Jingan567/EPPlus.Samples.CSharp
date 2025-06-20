using System.Text.RegularExpressions;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities;

public class WildCardValueMatcher : ValueMatcher
{
	protected override int CompareStringToString(string s1, string s2)
	{
		if (s1.Contains("*") || s1.Contains("?"))
		{
			string arg = Regex.Escape(s1);
			arg = $"^{arg}$";
			arg = arg.Replace("\\*", ".*");
			arg = arg.Replace("\\?", ".");
			if (Regex.IsMatch(s2, arg))
			{
				return 0;
			}
		}
		return base.CompareStringToString(s1, s2);
	}
}
