using System.Globalization;
using System.Xml;
using OfficeOpenXml.DataValidation.Formulas.Contracts;

namespace OfficeOpenXml.DataValidation.Formulas;

internal class ExcelDataValidationFormulaDecimal : ExcelDataValidationFormulaValue<double?>, IExcelDataValidationFormulaDecimal, IExcelDataValidationFormulaWithValue<double?>, IExcelDataValidationFormula
{
	public ExcelDataValidationFormulaDecimal(XmlNamespaceManager namespaceManager, XmlNode topNode, string formulaPath)
		: base(namespaceManager, topNode, formulaPath)
	{
		string xmlNodeString = GetXmlNodeString(formulaPath);
		if (!string.IsNullOrEmpty(xmlNodeString))
		{
			double result = 0.0;
			if (double.TryParse(xmlNodeString, NumberStyles.Any, CultureInfo.InvariantCulture, out result))
			{
				base.Value = result;
			}
			else
			{
				base.ExcelFormula = xmlNodeString;
			}
		}
	}

	protected override string GetValueAsString()
	{
		if (!base.Value.HasValue)
		{
			return string.Empty;
		}
		return base.Value.Value.ToString("R15", CultureInfo.InvariantCulture);
	}
}
