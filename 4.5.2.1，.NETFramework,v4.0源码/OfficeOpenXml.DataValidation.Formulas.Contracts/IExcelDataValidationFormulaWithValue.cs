namespace OfficeOpenXml.DataValidation.Formulas.Contracts;

public interface IExcelDataValidationFormulaWithValue<T> : IExcelDataValidationFormula
{
	T Value { get; set; }
}
