namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis;

public enum TokenType
{
	Operator,
	Negator,
	OpeningParenthesis,
	ClosingParenthesis,
	OpeningEnumerable,
	ClosingEnumerable,
	OpeningBracket,
	ClosingBracket,
	Enumerable,
	Comma,
	SemiColon,
	String,
	StringContent,
	WorksheetName,
	WorksheetNameContent,
	Integer,
	Boolean,
	Decimal,
	Percent,
	Function,
	ExcelAddress,
	NameValue,
	InvalidReference,
	NumericError,
	ValueDataTypeError,
	Null,
	Unrecognized,
	ExcelAddressR1C1
}
