using System.Xml;
using OfficeOpenXml.DataValidation.Contracts;
using OfficeOpenXml.DataValidation.Formulas;
using OfficeOpenXml.DataValidation.Formulas.Contracts;

namespace OfficeOpenXml.DataValidation;

public class ExcelDataValidationCustom : ExcelDataValidationWithFormula<IExcelDataValidationFormula>, IExcelDataValidationCustom, IExcelDataValidationWithFormula<IExcelDataValidationFormula>, IExcelDataValidation, IExcelDataValidationWithOperator
{
	internal ExcelDataValidationCustom(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType)
		: base(worksheet, address, validationType)
	{
		base.Formula = new ExcelDataValidationFormulaCustom(base.NameSpaceManager, base.TopNode, _formula1Path);
	}

	internal ExcelDataValidationCustom(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType, XmlNode itemElementNode)
		: base(worksheet, address, validationType, itemElementNode)
	{
		base.Formula = new ExcelDataValidationFormulaCustom(base.NameSpaceManager, base.TopNode, _formula1Path);
	}

	internal ExcelDataValidationCustom(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(worksheet, address, validationType, itemElementNode, namespaceManager)
	{
		base.Formula = new ExcelDataValidationFormulaCustom(base.NameSpaceManager, base.TopNode, _formula1Path);
	}
}
