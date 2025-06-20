using OfficeOpenXml.DataValidation.Contracts;

namespace OfficeOpenXml.DataValidation;

public interface IRangeDataValidation
{
	IExcelDataValidationAny AddAnyDataValidation();

	IExcelDataValidationInt AddIntegerDataValidation();

	IExcelDataValidationDecimal AddDecimalDataValidation();

	IExcelDataValidationDateTime AddDateTimeDataValidation();

	IExcelDataValidationList AddListDataValidation();

	IExcelDataValidationInt AddTextLengthDataValidation();

	IExcelDataValidationTime AddTimeDataValidation();

	IExcelDataValidationCustom AddCustomDataValidation();
}
