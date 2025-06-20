using System.Xml;
using OfficeOpenXml.DataValidation.Contracts;
using OfficeOpenXml.DataValidation.Formulas;
using OfficeOpenXml.DataValidation.Formulas.Contracts;

namespace OfficeOpenXml.DataValidation;

public class ExcelDataValidationTime : ExcelDataValidationWithFormula2<IExcelDataValidationFormulaTime>, IExcelDataValidationTime, IExcelDataValidationWithFormula2<IExcelDataValidationFormulaTime>, IExcelDataValidationWithFormula<IExcelDataValidationFormulaTime>, IExcelDataValidation, IExcelDataValidationWithOperator
{
	internal ExcelDataValidationTime(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType)
		: base(worksheet, address, validationType)
	{
		base.Formula = new ExcelDataValidationFormulaTime(base.NameSpaceManager, base.TopNode, _formula1Path);
		base.Formula2 = new ExcelDataValidationFormulaTime(base.NameSpaceManager, base.TopNode, _formula2Path);
	}

	internal ExcelDataValidationTime(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType, XmlNode itemElementNode)
		: base(worksheet, address, validationType, itemElementNode)
	{
		base.Formula = new ExcelDataValidationFormulaTime(base.NameSpaceManager, base.TopNode, _formula1Path);
		base.Formula2 = new ExcelDataValidationFormulaTime(base.NameSpaceManager, base.TopNode, _formula2Path);
	}

	internal ExcelDataValidationTime(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(worksheet, address, validationType, itemElementNode, namespaceManager)
	{
		base.Formula = new ExcelDataValidationFormulaTime(base.NameSpaceManager, base.TopNode, _formula1Path);
		base.Formula2 = new ExcelDataValidationFormulaTime(base.NameSpaceManager, base.TopNode, _formula2Path);
	}
}
