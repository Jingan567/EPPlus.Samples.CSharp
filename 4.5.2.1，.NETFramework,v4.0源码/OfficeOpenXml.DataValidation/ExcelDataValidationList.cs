using System.Xml;
using OfficeOpenXml.DataValidation.Contracts;
using OfficeOpenXml.DataValidation.Formulas;
using OfficeOpenXml.DataValidation.Formulas.Contracts;

namespace OfficeOpenXml.DataValidation;

public class ExcelDataValidationList : ExcelDataValidationWithFormula<IExcelDataValidationFormulaList>, IExcelDataValidationList, IExcelDataValidationWithFormula<IExcelDataValidationFormulaList>, IExcelDataValidation
{
	internal ExcelDataValidationList(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType)
		: base(worksheet, address, validationType)
	{
		base.Formula = new ExcelDataValidationFormulaList(base.NameSpaceManager, base.TopNode, _formula1Path);
	}

	internal ExcelDataValidationList(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType, XmlNode itemElementNode)
		: base(worksheet, address, validationType, itemElementNode)
	{
		base.Formula = new ExcelDataValidationFormulaList(base.NameSpaceManager, base.TopNode, _formula1Path);
	}

	internal ExcelDataValidationList(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(worksheet, address, validationType, itemElementNode, namespaceManager)
	{
		base.Formula = new ExcelDataValidationFormulaList(base.NameSpaceManager, base.TopNode, _formula1Path);
	}
}
