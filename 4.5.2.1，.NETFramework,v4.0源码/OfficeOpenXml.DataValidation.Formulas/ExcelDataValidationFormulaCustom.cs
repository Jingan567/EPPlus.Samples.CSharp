using System.Xml;
using OfficeOpenXml.DataValidation.Formulas.Contracts;

namespace OfficeOpenXml.DataValidation.Formulas;

internal class ExcelDataValidationFormulaCustom : ExcelDataValidationFormula, IExcelDataValidationFormula
{
	public ExcelDataValidationFormulaCustom(XmlNamespaceManager namespaceManager, XmlNode topNode, string formulaPath)
		: base(namespaceManager, topNode, formulaPath)
	{
		string xmlNodeString = GetXmlNodeString(formulaPath);
		if (!string.IsNullOrEmpty(xmlNodeString))
		{
			base.ExcelFormula = xmlNodeString;
		}
		base.State = FormulaState.Formula;
	}

	internal override string GetXmlValue()
	{
		return base.ExcelFormula;
	}

	protected override string GetValueAsString()
	{
		return base.ExcelFormula;
	}

	internal override void ResetValue()
	{
		base.ExcelFormula = null;
	}
}
