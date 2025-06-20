using System.Xml;
using OfficeOpenXml.DataValidation.Contracts;

namespace OfficeOpenXml.DataValidation;

public class ExcelDataValidationAny : ExcelDataValidation, IExcelDataValidationAny, IExcelDataValidation
{
	internal ExcelDataValidationAny(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType)
		: base(worksheet, address, validationType)
	{
	}

	internal ExcelDataValidationAny(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType, XmlNode itemElementNode)
		: base(worksheet, address, validationType, itemElementNode)
	{
	}

	internal ExcelDataValidationAny(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(worksheet, address, validationType, itemElementNode, namespaceManager)
	{
	}

	public override void Validate()
	{
	}
}
