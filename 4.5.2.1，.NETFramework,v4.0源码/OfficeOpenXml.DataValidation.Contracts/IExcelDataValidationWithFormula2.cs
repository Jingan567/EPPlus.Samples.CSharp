using OfficeOpenXml.DataValidation.Formulas.Contracts;

namespace OfficeOpenXml.DataValidation.Contracts;

public interface IExcelDataValidationWithFormula2<T> : IExcelDataValidationWithFormula<T>, IExcelDataValidation where T : IExcelDataValidationFormula
{
	T Formula2 { get; }
}
