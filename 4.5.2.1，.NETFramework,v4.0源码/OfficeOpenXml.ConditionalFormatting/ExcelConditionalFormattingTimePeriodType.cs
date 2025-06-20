using System;

namespace OfficeOpenXml.ConditionalFormatting;

internal static class ExcelConditionalFormattingTimePeriodType
{
	public static string GetAttributeByType(eExcelConditionalFormattingTimePeriodType type)
	{
		return type switch
		{
			eExcelConditionalFormattingTimePeriodType.Last7Days => "last7Days", 
			eExcelConditionalFormattingTimePeriodType.LastMonth => "lastMonth", 
			eExcelConditionalFormattingTimePeriodType.LastWeek => "lastWeek", 
			eExcelConditionalFormattingTimePeriodType.NextMonth => "nextMonth", 
			eExcelConditionalFormattingTimePeriodType.NextWeek => "nextWeek", 
			eExcelConditionalFormattingTimePeriodType.ThisMonth => "thisMonth", 
			eExcelConditionalFormattingTimePeriodType.ThisWeek => "thisWeek", 
			eExcelConditionalFormattingTimePeriodType.Today => "today", 
			eExcelConditionalFormattingTimePeriodType.Tomorrow => "tomorrow", 
			eExcelConditionalFormattingTimePeriodType.Yesterday => "yesterday", 
			_ => string.Empty, 
		};
	}

	public static eExcelConditionalFormattingTimePeriodType GetTypeByAttribute(string attribute)
	{
		return attribute switch
		{
			"last7Days" => eExcelConditionalFormattingTimePeriodType.Last7Days, 
			"lastMonth" => eExcelConditionalFormattingTimePeriodType.LastMonth, 
			"lastWeek" => eExcelConditionalFormattingTimePeriodType.LastWeek, 
			"nextMonth" => eExcelConditionalFormattingTimePeriodType.NextMonth, 
			"nextWeek" => eExcelConditionalFormattingTimePeriodType.NextWeek, 
			"thisMonth" => eExcelConditionalFormattingTimePeriodType.ThisMonth, 
			"thisWeek" => eExcelConditionalFormattingTimePeriodType.ThisWeek, 
			"today" => eExcelConditionalFormattingTimePeriodType.Today, 
			"tomorrow" => eExcelConditionalFormattingTimePeriodType.Tomorrow, 
			"yesterday" => eExcelConditionalFormattingTimePeriodType.Yesterday, 
			_ => throw new Exception("Unexistent eExcelConditionalFormattingTimePeriodType attribute in Conditional Formatting"), 
		};
	}
}
