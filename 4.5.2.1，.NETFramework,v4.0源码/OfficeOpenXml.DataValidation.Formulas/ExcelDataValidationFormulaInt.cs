using System.Globalization;
using System.Xml;
using OfficeOpenXml.DataValidation.Formulas.Contracts;

namespace OfficeOpenXml.DataValidation.Formulas;

internal class ExcelDataValidationFormulaInt : ExcelDataValidationFormulaValue<int?>, IExcelDataValidationFormulaInt, IExcelDataValidationFormulaWithValue<int?>, IExcelDataValidationFormula
{
	public ExcelDataValidationFormulaInt(XmlNamespaceManager namespaceManager, XmlNode topNode, string formulaPath)
		: base(namespaceManager, topNode, formulaPath)
	{
		string xmlNodeString = GetXmlNodeString(formulaPath);
		if (!string.IsNullOrEmpty(xmlNodeString))
		{
			int result = 0;
			if (int.TryParse(xmlNodeString, NumberStyles.Number, CultureInfo.InvariantCulture, out result))
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
		return base.Value.Value.ToString();
	}
}
