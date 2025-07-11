using System;
using System.Xml;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.ConditionalFormatting;

internal static class ExcelConditionalFormattingRuleFactory
{
	public static ExcelConditionalFormattingRule Create(eExcelConditionalFormattingRuleType type, ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode)
	{
		Require.Argument(type);
		Require.Argument(address).IsNotNull("address");
		Require.Argument(priority).IsInRange(0, int.MaxValue, "priority");
		Require.Argument(worksheet).IsNotNull("worksheet");
		return type switch
		{
			eExcelConditionalFormattingRuleType.AboveAverage => new ExcelConditionalFormattingAboveAverage(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.AboveOrEqualAverage => new ExcelConditionalFormattingAboveOrEqualAverage(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.BelowAverage => new ExcelConditionalFormattingBelowAverage(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.BelowOrEqualAverage => new ExcelConditionalFormattingBelowOrEqualAverage(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.AboveStdDev => new ExcelConditionalFormattingAboveStdDev(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.BelowStdDev => new ExcelConditionalFormattingBelowStdDev(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.Bottom => new ExcelConditionalFormattingBottom(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.BottomPercent => new ExcelConditionalFormattingBottomPercent(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.Top => new ExcelConditionalFormattingTop(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.TopPercent => new ExcelConditionalFormattingTopPercent(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.Last7Days => new ExcelConditionalFormattingLast7Days(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.LastMonth => new ExcelConditionalFormattingLastMonth(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.LastWeek => new ExcelConditionalFormattingLastWeek(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.NextMonth => new ExcelConditionalFormattingNextMonth(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.NextWeek => new ExcelConditionalFormattingNextWeek(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.ThisMonth => new ExcelConditionalFormattingThisMonth(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.ThisWeek => new ExcelConditionalFormattingThisWeek(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.Today => new ExcelConditionalFormattingToday(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.Tomorrow => new ExcelConditionalFormattingTomorrow(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.Yesterday => new ExcelConditionalFormattingYesterday(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.BeginsWith => new ExcelConditionalFormattingBeginsWith(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.Between => new ExcelConditionalFormattingBetween(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.ContainsBlanks => new ExcelConditionalFormattingContainsBlanks(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.ContainsErrors => new ExcelConditionalFormattingContainsErrors(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.ContainsText => new ExcelConditionalFormattingContainsText(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.DuplicateValues => new ExcelConditionalFormattingDuplicateValues(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.EndsWith => new ExcelConditionalFormattingEndsWith(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.Equal => new ExcelConditionalFormattingEqual(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.Expression => new ExcelConditionalFormattingExpression(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.GreaterThan => new ExcelConditionalFormattingGreaterThan(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.GreaterThanOrEqual => new ExcelConditionalFormattingGreaterThanOrEqual(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.LessThan => new ExcelConditionalFormattingLessThan(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.LessThanOrEqual => new ExcelConditionalFormattingLessThanOrEqual(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.NotBetween => new ExcelConditionalFormattingNotBetween(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.NotContainsBlanks => new ExcelConditionalFormattingNotContainsBlanks(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.NotContainsErrors => new ExcelConditionalFormattingNotContainsErrors(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.NotContainsText => new ExcelConditionalFormattingNotContainsText(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.NotEqual => new ExcelConditionalFormattingNotEqual(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.UniqueValues => new ExcelConditionalFormattingUniqueValues(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.ThreeColorScale => new ExcelConditionalFormattingThreeColorScale(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.TwoColorScale => new ExcelConditionalFormattingTwoColorScale(address, priority, worksheet, itemElementNode), 
			eExcelConditionalFormattingRuleType.ThreeIconSet => new ExcelConditionalFormattingThreeIconSet(address, priority, worksheet, itemElementNode, null), 
			eExcelConditionalFormattingRuleType.FourIconSet => new ExcelConditionalFormattingFourIconSet(address, priority, worksheet, itemElementNode, null), 
			eExcelConditionalFormattingRuleType.FiveIconSet => new ExcelConditionalFormattingFiveIconSet(address, priority, worksheet, itemElementNode, null), 
			eExcelConditionalFormattingRuleType.DataBar => new ExcelConditionalFormattingDataBar(eExcelConditionalFormattingRuleType.DataBar, address, priority, worksheet, itemElementNode, null), 
			_ => throw new InvalidOperationException($"Non supported conditionalFormattingType: {type.ToString()}"), 
		};
	}
}
