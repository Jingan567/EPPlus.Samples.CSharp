using System.Xml;
using OfficeOpenXml.DataValidation.Formulas.Contracts;

namespace OfficeOpenXml.DataValidation;

public class ExcelDataValidationWithFormula2<T> : ExcelDataValidationWithFormula<T> where T : IExcelDataValidationFormula
{
	public T Formula2 { get; protected set; }

	internal ExcelDataValidationWithFormula2(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType)
		: this(worksheet, address, validationType, (XmlNode)null)
	{
	}

	internal ExcelDataValidationWithFormula2(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType, XmlNode itemElementNode)
		: base(worksheet, address, validationType, itemElementNode)
	{
	}

	internal ExcelDataValidationWithFormula2(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(worksheet, address, validationType, itemElementNode, namespaceManager)
	{
	}
}
