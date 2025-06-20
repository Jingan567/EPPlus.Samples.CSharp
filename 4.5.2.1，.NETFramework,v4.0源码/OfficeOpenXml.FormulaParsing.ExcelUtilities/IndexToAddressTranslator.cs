using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities;

public class IndexToAddressTranslator
{
	private readonly ExcelDataProvider _excelDataProvider;

	private readonly ExcelReferenceType _excelReferenceType;

	public IndexToAddressTranslator(ExcelDataProvider excelDataProvider)
		: this(excelDataProvider, ExcelReferenceType.AbsoluteRowAndColumn)
	{
	}

	public IndexToAddressTranslator(ExcelDataProvider excelDataProvider, ExcelReferenceType referenceType)
	{
		Require.That(excelDataProvider).Named("excelDataProvider").IsNotNull();
		_excelDataProvider = excelDataProvider;
		_excelReferenceType = referenceType;
	}

	protected internal static string GetColumnLetter(int iColumnNumber, bool fixedCol)
	{
		if (iColumnNumber < 1)
		{
			return "#REF!";
		}
		string text = "";
		do
		{
			text = (char)(65 + (iColumnNumber - 1) % 26) + text;
			iColumnNumber = (iColumnNumber - (iColumnNumber - 1) % 26) / 26;
		}
		while (iColumnNumber > 0);
		if (!fixedCol)
		{
			return text;
		}
		return "$" + text;
	}

	public string ToAddress(int col, int row)
	{
		bool fixedCol = _excelReferenceType == ExcelReferenceType.AbsoluteRowAndColumn || _excelReferenceType == ExcelReferenceType.RelativeRowAbsolutColumn;
		return GetColumnLetter(col, fixedCol) + GetRowNumber(row);
	}

	private string GetRowNumber(int rowNo)
	{
		string text = ((rowNo < _excelDataProvider.ExcelMaxRows) ? rowNo.ToString() : string.Empty);
		if (!string.IsNullOrEmpty(text))
		{
			ExcelReferenceType excelReferenceType = _excelReferenceType;
			if ((uint)(excelReferenceType - 1) <= 1u)
			{
				return "$" + text;
			}
			return text;
		}
		return text;
	}
}
