using System;
using System.Globalization;
using System.Xml;
using OfficeOpenXml.DataValidation.Formulas.Contracts;

namespace OfficeOpenXml.DataValidation.Formulas;

internal class ExcelDataValidationFormulaDateTime : ExcelDataValidationFormulaValue<DateTime?>, IExcelDataValidationFormulaDateTime, IExcelDataValidationFormulaWithValue<DateTime?>, IExcelDataValidationFormula
{
	public ExcelDataValidationFormulaDateTime(XmlNamespaceManager namespaceManager, XmlNode topNode, string formulaPath)
		: base(namespaceManager, topNode, formulaPath)
	{
		string xmlNodeString = GetXmlNodeString(formulaPath);
		if (!string.IsNullOrEmpty(xmlNodeString))
		{
			double result = 0.0;
			if (double.TryParse(xmlNodeString, NumberStyles.Any, CultureInfo.InvariantCulture, out result))
			{
				base.Value = DateTime.FromOADate(result);
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
		return base.Value.Value.ToOADate().ToString(CultureInfo.InvariantCulture);
	}
}
