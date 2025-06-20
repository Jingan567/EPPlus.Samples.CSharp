using System.Text.RegularExpressions;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities;

public class AddressTranslator
{
	public enum RangeCalculationBehaviour
	{
		FirstPart,
		LastPart
	}

	private readonly ExcelDataProvider _excelDataProvider;

	public AddressTranslator(ExcelDataProvider excelDataProvider)
	{
		OfficeOpenXml.FormulaParsing.Utilities.Require.That(excelDataProvider).Named("excelDataProvider").IsNotNull();
		_excelDataProvider = excelDataProvider;
	}

	public virtual void ToColAndRow(string address, out int col, out int row)
	{
		ToColAndRow(address, out col, out row, RangeCalculationBehaviour.FirstPart);
	}

	public virtual void ToColAndRow(string address, out int col, out int row, RangeCalculationBehaviour behaviour)
	{
		address = ConvertUtil._invariantTextInfo.ToUpper(address);
		string alphaPart = GetAlphaPart(address);
		col = 0;
		int num = 26;
		for (int i = 0; i < alphaPart.Length; i++)
		{
			int num2 = alphaPart.Length - i - 1;
			int numericAlphaValue = GetNumericAlphaValue(alphaPart[i]);
			col += num * num2 * numericAlphaValue;
			if (num2 == 0)
			{
				col += numericAlphaValue;
			}
		}
		row = GetIntPart(address) ?? GetRowIndexByBehaviour(behaviour);
	}

	private int GetRowIndexByBehaviour(RangeCalculationBehaviour behaviour)
	{
		if (behaviour == RangeCalculationBehaviour.FirstPart)
		{
			return 1;
		}
		return _excelDataProvider.ExcelMaxRows;
	}

	private int GetNumericAlphaValue(char c)
	{
		return c - 64;
	}

	private string GetAlphaPart(string address)
	{
		return Regex.Match(address, "[A-Z]+").Value;
	}

	private int? GetIntPart(string address)
	{
		if (Regex.IsMatch(address, "[0-9]+"))
		{
			return int.Parse(Regex.Match(address, "[0-9]+").Value);
		}
		return null;
	}
}
