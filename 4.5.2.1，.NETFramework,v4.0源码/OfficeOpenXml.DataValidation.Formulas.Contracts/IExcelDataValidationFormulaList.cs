using System.Collections.Generic;

namespace OfficeOpenXml.DataValidation.Formulas.Contracts;

public interface IExcelDataValidationFormulaList : IExcelDataValidationFormula
{
	IList<string> Values { get; }
}
