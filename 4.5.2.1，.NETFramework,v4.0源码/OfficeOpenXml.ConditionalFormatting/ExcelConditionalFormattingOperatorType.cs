using System;

namespace OfficeOpenXml.ConditionalFormatting;

internal static class ExcelConditionalFormattingOperatorType
{
	internal static string GetAttributeByType(eExcelConditionalFormattingOperatorType type)
	{
		return type switch
		{
			eExcelConditionalFormattingOperatorType.BeginsWith => "beginsWith", 
			eExcelConditionalFormattingOperatorType.Between => "between", 
			eExcelConditionalFormattingOperatorType.ContainsText => "containsText", 
			eExcelConditionalFormattingOperatorType.EndsWith => "endsWith", 
			eExcelConditionalFormattingOperatorType.Equal => "equal", 
			eExcelConditionalFormattingOperatorType.GreaterThan => "greaterThan", 
			eExcelConditionalFormattingOperatorType.GreaterThanOrEqual => "greaterThanOrEqual", 
			eExcelConditionalFormattingOperatorType.LessThan => "lessThan", 
			eExcelConditionalFormattingOperatorType.LessThanOrEqual => "lessThanOrEqual", 
			eExcelConditionalFormattingOperatorType.NotBetween => "notBetween", 
			eExcelConditionalFormattingOperatorType.NotContains => "notContains", 
			eExcelConditionalFormattingOperatorType.NotEqual => "notEqual", 
			_ => string.Empty, 
		};
	}

	internal static eExcelConditionalFormattingOperatorType GetTypeByAttribute(string attribute)
	{
		return attribute switch
		{
			"beginsWith" => eExcelConditionalFormattingOperatorType.BeginsWith, 
			"between" => eExcelConditionalFormattingOperatorType.Between, 
			"containsText" => eExcelConditionalFormattingOperatorType.ContainsText, 
			"endsWith" => eExcelConditionalFormattingOperatorType.EndsWith, 
			"equal" => eExcelConditionalFormattingOperatorType.Equal, 
			"greaterThan" => eExcelConditionalFormattingOperatorType.GreaterThan, 
			"greaterThanOrEqual" => eExcelConditionalFormattingOperatorType.GreaterThanOrEqual, 
			"lessThan" => eExcelConditionalFormattingOperatorType.LessThan, 
			"lessThanOrEqual" => eExcelConditionalFormattingOperatorType.LessThanOrEqual, 
			"notBetween" => eExcelConditionalFormattingOperatorType.NotBetween, 
			"notContains" => eExcelConditionalFormattingOperatorType.NotContains, 
			"notEqual" => eExcelConditionalFormattingOperatorType.NotEqual, 
			_ => throw new Exception("Unexistent eExcelConditionalFormattingOperatorType attribute in Conditional Formatting"), 
		};
	}
}
