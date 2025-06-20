using System.Xml;

namespace OfficeOpenXml.DataValidation.Formulas;

internal abstract class ExcelDataValidationFormulaValue<T> : ExcelDataValidationFormula
{
	private T _value;

	public T Value
	{
		get
		{
			return _value;
		}
		set
		{
			base.State = FormulaState.Value;
			_value = value;
			SetXmlNodeString(base.FormulaPath, GetValueAsString());
		}
	}

	public ExcelDataValidationFormulaValue(XmlNamespaceManager namespaceManager, XmlNode topNode, string formulaPath)
		: base(namespaceManager, topNode, formulaPath)
	{
	}

	internal override void ResetValue()
	{
		Value = default(T);
	}
}
