using OfficeOpenXml.DataValidation.Formulas.Contracts;

namespace OfficeOpenXml.DataValidation.Contracts;

public interface IExcelDataValidationInt : IExcelDataValidationWithFormula2<IExcelDataValidationFormulaInt>, IExcelDataValidationWithFormula<IExcelDataValidationFormulaInt>, IExcelDataValidation, IExcelDataValidationWithOperator
{
}
