using System;
using System.Xml;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.DataValidation.Formulas;

internal abstract class ExcelDataValidationFormula : XmlHelper
{
	private string _formula;

	protected string FormulaPath { get; private set; }

	protected FormulaState State { get; set; }

	public string ExcelFormula
	{
		get
		{
			return _formula;
		}
		set
		{
			if (!string.IsNullOrEmpty(value))
			{
				ResetValue();
				State = FormulaState.Formula;
			}
			if (value != null && value.Length > 255)
			{
				throw new InvalidOperationException("The length of a DataValidation formula cannot exceed 255 characters");
			}
			_formula = value;
			SetXmlNodeString(FormulaPath, value);
		}
	}

	public ExcelDataValidationFormula(XmlNamespaceManager namespaceManager, XmlNode topNode, string formulaPath)
		: base(namespaceManager, topNode)
	{
		Require.Argument(formulaPath).IsNotNullOrEmpty("formulaPath");
		FormulaPath = formulaPath;
	}

	internal abstract void ResetValue();

	internal virtual string GetXmlValue()
	{
		if (State == FormulaState.Formula)
		{
			return ExcelFormula;
		}
		return GetValueAsString();
	}

	protected abstract string GetValueAsString();
}
