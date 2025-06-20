using System.Xml;
using OfficeOpenXml.DataValidation.Contracts;
using OfficeOpenXml.DataValidation.Formulas;
using OfficeOpenXml.DataValidation.Formulas.Contracts;

namespace OfficeOpenXml.DataValidation;

public class ExcelDataValidationDateTime : ExcelDataValidationWithFormula2<IExcelDataValidationFormulaDateTime>, IExcelDataValidationDateTime, IExcelDataValidationWithFormula2<IExcelDataValidationFormulaDateTime>, IExcelDataValidationWithFormula<IExcelDataValidationFormulaDateTime>, IExcelDataValidation, IExcelDataValidationWithOperator
{
	internal ExcelDataValidationDateTime(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType)
		: base(worksheet, address, validationType)
	{
		base.Formula = new ExcelDataValidationFormulaDateTime(base.NameSpaceManager, base.TopNode, _formula1Path);
		base.Formula2 = new ExcelDataValidationFormulaDateTime(base.NameSpaceManager, base.TopNode, _formula2Path);
	}

	internal ExcelDataValidationDateTime(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType, XmlNode itemElementNode)
		: base(worksheet, address, validationType, itemElementNode)
	{
		base.Formula = new ExcelDataValidationFormulaDateTime(base.NameSpaceManager, base.TopNode, _formula1Path);
		base.Formula2 = new ExcelDataValidationFormulaDateTime(base.NameSpaceManager, base.TopNode, _formula2Path);
	}

	internal ExcelDataValidationDateTime(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(worksheet, address, validationType, itemElementNode, namespaceManager)
	{
		base.Formula = new ExcelDataValidationFormulaDateTime(base.NameSpaceManager, base.TopNode, _formula1Path);
		base.Formula2 = new ExcelDataValidationFormulaDateTime(base.NameSpaceManager, base.TopNode, _formula2Path);
	}
}
