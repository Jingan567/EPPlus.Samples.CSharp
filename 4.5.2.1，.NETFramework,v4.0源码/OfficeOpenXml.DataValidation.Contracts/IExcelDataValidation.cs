namespace OfficeOpenXml.DataValidation.Contracts;

public interface IExcelDataValidation
{
	ExcelAddress Address { get; }

	ExcelDataValidationType ValidationType { get; }

	ExcelDataValidationWarningStyle ErrorStyle { get; set; }

	bool? AllowBlank { get; set; }

	bool? ShowInputMessage { get; set; }

	bool? ShowErrorMessage { get; set; }

	string ErrorTitle { get; set; }

	string Error { get; set; }

	string PromptTitle { get; set; }

	string Prompt { get; set; }

	bool AllowsOperator { get; }

	void Validate();
}
