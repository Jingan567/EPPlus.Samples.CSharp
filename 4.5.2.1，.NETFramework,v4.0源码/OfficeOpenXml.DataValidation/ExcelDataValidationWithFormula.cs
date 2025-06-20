using System;
using System.Xml;
using OfficeOpenXml.DataValidation.Formulas.Contracts;

namespace OfficeOpenXml.DataValidation;

public class ExcelDataValidationWithFormula<T> : ExcelDataValidation where T : IExcelDataValidationFormula
{
	public T Formula { get; protected set; }

	internal ExcelDataValidationWithFormula(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType)
		: this(worksheet, address, validationType, (XmlNode)null)
	{
	}

	internal ExcelDataValidationWithFormula(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType, XmlNode itemElementNode)
		: base(worksheet, address, validationType, itemElementNode)
	{
	}

	internal ExcelDataValidationWithFormula(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(worksheet, address, validationType, itemElementNode, namespaceManager)
	{
	}

	public override void Validate()
	{
		base.Validate();
		if ((base.Operator == ExcelDataValidationOperator.between || base.Operator == ExcelDataValidationOperator.notBetween) && string.IsNullOrEmpty(base.Formula2Internal))
		{
			throw new InvalidOperationException("Validation of " + base.Address.Address + " failed: Formula2 must be set if operator is 'between' or 'notBetween'");
		}
	}
}
