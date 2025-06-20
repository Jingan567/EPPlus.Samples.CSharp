using System.Xml;
using OfficeOpenXml.DataValidation.Contracts;
using OfficeOpenXml.DataValidation.Formulas;
using OfficeOpenXml.DataValidation.Formulas.Contracts;

namespace OfficeOpenXml.DataValidation;

public class ExcelDataValidationDecimal : ExcelDataValidationWithFormula2<IExcelDataValidationFormulaDecimal>, IExcelDataValidationDecimal, IExcelDataValidationWithFormula2<IExcelDataValidationFormulaDecimal>, IExcelDataValidationWithFormula<IExcelDataValidationFormulaDecimal>, IExcelDataValidation, IExcelDataValidationWithOperator
{
	internal ExcelDataValidationDecimal(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType)
		: base(worksheet, address, validationType)
	{
		base.Formula = new ExcelDataValidationFormulaDecimal(base.NameSpaceManager, base.TopNode, _formula1Path);
		base.Formula2 = new ExcelDataValidationFormulaDecimal(base.NameSpaceManager, base.TopNode, _formula2Path);
	}

	internal ExcelDataValidationDecimal(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType, XmlNode itemElementNode)
		: base(worksheet, address, validationType, itemElementNode)
	{
		base.Formula = new ExcelDataValidationFormulaDecimal(base.NameSpaceManager, base.TopNode, _formula1Path);
		base.Formula2 = new ExcelDataValidationFormulaDecimal(base.NameSpaceManager, base.TopNode, _formula2Path);
	}

	internal ExcelDataValidationDecimal(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(worksheet, address, validationType, itemElementNode, namespaceManager)
	{
		base.Formula = new ExcelDataValidationFormulaDecimal(base.NameSpaceManager, base.TopNode, _formula1Path);
		base.Formula2 = new ExcelDataValidationFormulaDecimal(base.NameSpaceManager, base.TopNode, _formula2Path);
	}
}
