using System.Xml;
using OfficeOpenXml.DataValidation.Contracts;
using OfficeOpenXml.DataValidation.Formulas;
using OfficeOpenXml.DataValidation.Formulas.Contracts;

namespace OfficeOpenXml.DataValidation;

public class ExcelDataValidationInt : ExcelDataValidationWithFormula2<IExcelDataValidationFormulaInt>, IExcelDataValidationInt, IExcelDataValidationWithFormula2<IExcelDataValidationFormulaInt>, IExcelDataValidationWithFormula<IExcelDataValidationFormulaInt>, IExcelDataValidation, IExcelDataValidationWithOperator
{
	internal ExcelDataValidationInt(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType)
		: base(worksheet, address, validationType)
	{
		base.Formula = new ExcelDataValidationFormulaInt(worksheet.NameSpaceManager, base.TopNode, _formula1Path);
		base.Formula2 = new ExcelDataValidationFormulaInt(worksheet.NameSpaceManager, base.TopNode, _formula2Path);
	}

	internal ExcelDataValidationInt(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType, XmlNode itemElementNode)
		: base(worksheet, address, validationType, itemElementNode)
	{
		base.Formula = new ExcelDataValidationFormulaInt(worksheet.NameSpaceManager, base.TopNode, _formula1Path);
		base.Formula2 = new ExcelDataValidationFormulaInt(worksheet.NameSpaceManager, base.TopNode, _formula2Path);
	}

	internal ExcelDataValidationInt(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(worksheet, address, validationType, itemElementNode, namespaceManager)
	{
		base.Formula = new ExcelDataValidationFormulaInt(base.NameSpaceManager, base.TopNode, _formula1Path);
		base.Formula2 = new ExcelDataValidationFormulaInt(base.NameSpaceManager, base.TopNode, _formula2Path);
	}
}
