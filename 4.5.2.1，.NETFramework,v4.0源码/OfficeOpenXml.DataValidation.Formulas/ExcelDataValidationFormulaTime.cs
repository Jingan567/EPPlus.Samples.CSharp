using System;
using System.Globalization;
using System.Xml;
using OfficeOpenXml.DataValidation.Formulas.Contracts;

namespace OfficeOpenXml.DataValidation.Formulas;

internal class ExcelDataValidationFormulaTime : ExcelDataValidationFormulaValue<ExcelTime>, IExcelDataValidationFormulaTime, IExcelDataValidationFormulaWithValue<ExcelTime>, IExcelDataValidationFormula
{
	public ExcelDataValidationFormulaTime(XmlNamespaceManager namespaceManager, XmlNode topNode, string formulaPath)
		: base(namespaceManager, topNode, formulaPath)
	{
		string xmlNodeString = GetXmlNodeString(formulaPath);
		if (!string.IsNullOrEmpty(xmlNodeString))
		{
			decimal result = default(decimal);
			if (decimal.TryParse(xmlNodeString, NumberStyles.Any, CultureInfo.InvariantCulture, out result))
			{
				base.Value = new ExcelTime(result);
			}
			else
			{
				base.Value = new ExcelTime();
				base.ExcelFormula = xmlNodeString;
			}
		}
		else
		{
			base.Value = new ExcelTime();
		}
		base.Value.TimeChanged += Value_TimeChanged;
	}

	private void Value_TimeChanged(object sender, EventArgs e)
	{
		SetXmlNodeString(base.FormulaPath, base.Value.ToExcelString());
	}

	protected override string GetValueAsString()
	{
		if (base.State == FormulaState.Value)
		{
			return base.Value.ToExcelString();
		}
		return string.Empty;
	}

	internal override void ResetValue()
	{
		base.Value = new ExcelTime();
	}
}
