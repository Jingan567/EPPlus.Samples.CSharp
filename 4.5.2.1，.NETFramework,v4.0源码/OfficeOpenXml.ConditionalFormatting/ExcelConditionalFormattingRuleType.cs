using System;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting;

internal static class ExcelConditionalFormattingRuleType
{
	internal static eExcelConditionalFormattingRuleType GetTypeByAttrbiute(string attribute, XmlNode topNode, XmlNamespaceManager nameSpaceManager)
	{
		return attribute switch
		{
			"aboveAverage" => GetAboveAverageType(topNode, nameSpaceManager), 
			"top10" => GetTop10Type(topNode, nameSpaceManager), 
			"timePeriod" => GetTimePeriodType(topNode, nameSpaceManager), 
			"cellIs" => GetCellIs((XmlElement)topNode), 
			"beginsWith" => eExcelConditionalFormattingRuleType.BeginsWith, 
			"containsBlanks" => eExcelConditionalFormattingRuleType.ContainsBlanks, 
			"containsErrors" => eExcelConditionalFormattingRuleType.ContainsErrors, 
			"containsText" => eExcelConditionalFormattingRuleType.ContainsText, 
			"duplicateValues" => eExcelConditionalFormattingRuleType.DuplicateValues, 
			"endsWith" => eExcelConditionalFormattingRuleType.EndsWith, 
			"expression" => eExcelConditionalFormattingRuleType.Expression, 
			"notContainsBlanks" => eExcelConditionalFormattingRuleType.NotContainsBlanks, 
			"notContainsErrors" => eExcelConditionalFormattingRuleType.NotContainsErrors, 
			"notContainsText" => eExcelConditionalFormattingRuleType.NotContainsText, 
			"uniqueValues" => eExcelConditionalFormattingRuleType.UniqueValues, 
			"colorScale" => GetColorScaleType(topNode, nameSpaceManager), 
			"iconSet" => GetIconSetType(topNode, nameSpaceManager), 
			"dataBar" => eExcelConditionalFormattingRuleType.DataBar, 
			_ => throw new Exception("Unexpected eExcelConditionalFormattingRuleType attribute in Conditional Formatting Rule"), 
		};
	}

	private static eExcelConditionalFormattingRuleType GetCellIs(XmlElement node)
	{
		return node.GetAttribute("operator") switch
		{
			"beginsWith" => eExcelConditionalFormattingRuleType.BeginsWith, 
			"between" => eExcelConditionalFormattingRuleType.Between, 
			"containsText" => eExcelConditionalFormattingRuleType.ContainsText, 
			"endsWith" => eExcelConditionalFormattingRuleType.EndsWith, 
			"equal" => eExcelConditionalFormattingRuleType.Equal, 
			"greaterThan" => eExcelConditionalFormattingRuleType.GreaterThan, 
			"greaterThanOrEqual" => eExcelConditionalFormattingRuleType.GreaterThanOrEqual, 
			"lessThan" => eExcelConditionalFormattingRuleType.LessThan, 
			"lessThanOrEqual" => eExcelConditionalFormattingRuleType.LessThanOrEqual, 
			"notBetween" => eExcelConditionalFormattingRuleType.NotBetween, 
			"notContains" => eExcelConditionalFormattingRuleType.NotContains, 
			"notEqual" => eExcelConditionalFormattingRuleType.NotEqual, 
			_ => throw new Exception("Unexistent eExcelConditionalFormattingOperatorType attribute in Conditional Formatting"), 
		};
	}

	private static eExcelConditionalFormattingRuleType GetIconSetType(XmlNode topNode, XmlNamespaceManager nameSpaceManager)
	{
		XmlNode xmlNode = topNode.SelectSingleNode("d:iconSet/@iconSet", nameSpaceManager);
		if (xmlNode == null)
		{
			return eExcelConditionalFormattingRuleType.ThreeIconSet;
		}
		string value = xmlNode.Value;
		if (value[0] == '3')
		{
			return eExcelConditionalFormattingRuleType.ThreeIconSet;
		}
		if (value[0] == '4')
		{
			return eExcelConditionalFormattingRuleType.FourIconSet;
		}
		return eExcelConditionalFormattingRuleType.FiveIconSet;
	}

	internal static eExcelConditionalFormattingRuleType GetColorScaleType(XmlNode topNode, XmlNamespaceManager nameSpaceManager)
	{
		XmlNodeList xmlNodeList = topNode.SelectNodes(string.Format("{0}/{1}", "d:colorScale", "d:cfvo"), nameSpaceManager);
		XmlNodeList xmlNodeList2 = topNode.SelectNodes(string.Format("{0}/{1}", "d:colorScale", "d:color"), nameSpaceManager);
		if (xmlNodeList == null || xmlNodeList.Count < 2 || xmlNodeList.Count > 3 || xmlNodeList2 == null || xmlNodeList2.Count < 2 || xmlNodeList2.Count > 3 || xmlNodeList.Count != xmlNodeList2.Count)
		{
			throw new Exception("Wrong number of 'cfvo'/'color' nodes in Conditional Formatting Rule");
		}
		if (xmlNodeList.Count != 2)
		{
			return eExcelConditionalFormattingRuleType.ThreeColorScale;
		}
		return eExcelConditionalFormattingRuleType.TwoColorScale;
	}

	internal static eExcelConditionalFormattingRuleType GetAboveAverageType(XmlNode topNode, XmlNamespaceManager nameSpaceManager)
	{
		int? attributeIntNullable = ExcelConditionalFormattingHelper.GetAttributeIntNullable(topNode, "stdDev");
		if (attributeIntNullable > 0)
		{
			return eExcelConditionalFormattingRuleType.AboveStdDev;
		}
		if (attributeIntNullable < 0)
		{
			return eExcelConditionalFormattingRuleType.BelowStdDev;
		}
		bool? attributeBoolNullable = ExcelConditionalFormattingHelper.GetAttributeBoolNullable(topNode, "aboveAverage");
		bool? attributeBoolNullable2 = ExcelConditionalFormattingHelper.GetAttributeBoolNullable(topNode, "equalAverage");
		if (!attributeBoolNullable.HasValue || attributeBoolNullable == true)
		{
			if (attributeBoolNullable2 == true)
			{
				return eExcelConditionalFormattingRuleType.AboveOrEqualAverage;
			}
			return eExcelConditionalFormattingRuleType.AboveAverage;
		}
		if (attributeBoolNullable2 == true)
		{
			return eExcelConditionalFormattingRuleType.BelowOrEqualAverage;
		}
		return eExcelConditionalFormattingRuleType.BelowAverage;
	}

	public static eExcelConditionalFormattingRuleType GetTop10Type(XmlNode topNode, XmlNamespaceManager nameSpaceManager)
	{
		bool? attributeBoolNullable = ExcelConditionalFormattingHelper.GetAttributeBoolNullable(topNode, "bottom");
		bool? attributeBoolNullable2 = ExcelConditionalFormattingHelper.GetAttributeBoolNullable(topNode, "percent");
		if (attributeBoolNullable == true)
		{
			if (attributeBoolNullable2 == true)
			{
				return eExcelConditionalFormattingRuleType.BottomPercent;
			}
			return eExcelConditionalFormattingRuleType.Bottom;
		}
		if (attributeBoolNullable2 == true)
		{
			return eExcelConditionalFormattingRuleType.TopPercent;
		}
		return eExcelConditionalFormattingRuleType.Top;
	}

	public static eExcelConditionalFormattingRuleType GetTimePeriodType(XmlNode topNode, XmlNamespaceManager nameSpaceManager)
	{
		return ExcelConditionalFormattingTimePeriodType.GetTypeByAttribute(ExcelConditionalFormattingHelper.GetAttributeString(topNode, "timePeriod")) switch
		{
			eExcelConditionalFormattingTimePeriodType.Last7Days => eExcelConditionalFormattingRuleType.Last7Days, 
			eExcelConditionalFormattingTimePeriodType.LastMonth => eExcelConditionalFormattingRuleType.LastMonth, 
			eExcelConditionalFormattingTimePeriodType.LastWeek => eExcelConditionalFormattingRuleType.LastWeek, 
			eExcelConditionalFormattingTimePeriodType.NextMonth => eExcelConditionalFormattingRuleType.NextMonth, 
			eExcelConditionalFormattingTimePeriodType.NextWeek => eExcelConditionalFormattingRuleType.NextWeek, 
			eExcelConditionalFormattingTimePeriodType.ThisMonth => eExcelConditionalFormattingRuleType.ThisMonth, 
			eExcelConditionalFormattingTimePeriodType.ThisWeek => eExcelConditionalFormattingRuleType.ThisWeek, 
			eExcelConditionalFormattingTimePeriodType.Today => eExcelConditionalFormattingRuleType.Today, 
			eExcelConditionalFormattingTimePeriodType.Tomorrow => eExcelConditionalFormattingRuleType.Tomorrow, 
			eExcelConditionalFormattingTimePeriodType.Yesterday => eExcelConditionalFormattingRuleType.Yesterday, 
			_ => throw new Exception("Unexistent eExcelConditionalFormattingTimePeriodType attribute in Conditional Formatting"), 
		};
	}

	public static string GetAttributeByType(eExcelConditionalFormattingRuleType type)
	{
		switch (type)
		{
		case eExcelConditionalFormattingRuleType.AboveAverage:
		case eExcelConditionalFormattingRuleType.AboveOrEqualAverage:
		case eExcelConditionalFormattingRuleType.BelowAverage:
		case eExcelConditionalFormattingRuleType.BelowOrEqualAverage:
		case eExcelConditionalFormattingRuleType.AboveStdDev:
		case eExcelConditionalFormattingRuleType.BelowStdDev:
			return "aboveAverage";
		case eExcelConditionalFormattingRuleType.Bottom:
		case eExcelConditionalFormattingRuleType.BottomPercent:
		case eExcelConditionalFormattingRuleType.Top:
		case eExcelConditionalFormattingRuleType.TopPercent:
			return "top10";
		case eExcelConditionalFormattingRuleType.Last7Days:
		case eExcelConditionalFormattingRuleType.LastMonth:
		case eExcelConditionalFormattingRuleType.LastWeek:
		case eExcelConditionalFormattingRuleType.NextMonth:
		case eExcelConditionalFormattingRuleType.NextWeek:
		case eExcelConditionalFormattingRuleType.ThisMonth:
		case eExcelConditionalFormattingRuleType.ThisWeek:
		case eExcelConditionalFormattingRuleType.Today:
		case eExcelConditionalFormattingRuleType.Tomorrow:
		case eExcelConditionalFormattingRuleType.Yesterday:
			return "timePeriod";
		case eExcelConditionalFormattingRuleType.Between:
		case eExcelConditionalFormattingRuleType.Equal:
		case eExcelConditionalFormattingRuleType.GreaterThan:
		case eExcelConditionalFormattingRuleType.GreaterThanOrEqual:
		case eExcelConditionalFormattingRuleType.LessThan:
		case eExcelConditionalFormattingRuleType.LessThanOrEqual:
		case eExcelConditionalFormattingRuleType.NotBetween:
		case eExcelConditionalFormattingRuleType.NotEqual:
			return "cellIs";
		case eExcelConditionalFormattingRuleType.ThreeIconSet:
		case eExcelConditionalFormattingRuleType.FourIconSet:
		case eExcelConditionalFormattingRuleType.FiveIconSet:
			return "iconSet";
		case eExcelConditionalFormattingRuleType.ThreeColorScale:
		case eExcelConditionalFormattingRuleType.TwoColorScale:
			return "colorScale";
		case eExcelConditionalFormattingRuleType.BeginsWith:
			return "beginsWith";
		case eExcelConditionalFormattingRuleType.ContainsBlanks:
			return "containsBlanks";
		case eExcelConditionalFormattingRuleType.ContainsErrors:
			return "containsErrors";
		case eExcelConditionalFormattingRuleType.ContainsText:
			return "containsText";
		case eExcelConditionalFormattingRuleType.DuplicateValues:
			return "duplicateValues";
		case eExcelConditionalFormattingRuleType.EndsWith:
			return "endsWith";
		case eExcelConditionalFormattingRuleType.Expression:
			return "expression";
		case eExcelConditionalFormattingRuleType.NotContainsBlanks:
			return "notContainsBlanks";
		case eExcelConditionalFormattingRuleType.NotContainsErrors:
			return "notContainsErrors";
		case eExcelConditionalFormattingRuleType.NotContainsText:
			return "notContainsText";
		case eExcelConditionalFormattingRuleType.UniqueValues:
			return "uniqueValues";
		case eExcelConditionalFormattingRuleType.DataBar:
			return "dataBar";
		default:
			throw new Exception("Missing eExcelConditionalFormattingRuleType Type in Conditional Formatting");
		}
	}

	public static string GetCfvoParentPathByType(eExcelConditionalFormattingRuleType type)
	{
		switch (type)
		{
		case eExcelConditionalFormattingRuleType.ThreeColorScale:
		case eExcelConditionalFormattingRuleType.TwoColorScale:
			return "d:colorScale";
		case eExcelConditionalFormattingRuleType.ThreeIconSet:
		case eExcelConditionalFormattingRuleType.FourIconSet:
		case eExcelConditionalFormattingRuleType.FiveIconSet:
			return "iconSet";
		case eExcelConditionalFormattingRuleType.DataBar:
			return "dataBar";
		default:
			throw new Exception("Missing eExcelConditionalFormattingRuleType Type in Conditional Formatting");
		}
	}
}
