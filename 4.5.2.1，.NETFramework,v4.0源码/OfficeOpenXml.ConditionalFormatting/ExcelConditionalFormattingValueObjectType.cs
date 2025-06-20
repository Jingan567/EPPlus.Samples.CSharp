using System;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting;

internal static class ExcelConditionalFormattingValueObjectType
{
	internal static int GetOrderByPosition(eExcelConditionalFormattingValueObjectPosition position, eExcelConditionalFormattingRuleType ruleType)
	{
		switch (position)
		{
		case eExcelConditionalFormattingValueObjectPosition.Low:
			return 1;
		case eExcelConditionalFormattingValueObjectPosition.Middle:
			return 2;
		case eExcelConditionalFormattingValueObjectPosition.High:
			if (ruleType == eExcelConditionalFormattingRuleType.TwoColorScale)
			{
				return 2;
			}
			return 3;
		default:
			return 0;
		}
	}

	public static eExcelConditionalFormattingValueObjectType GetTypeByAttrbiute(string attribute)
	{
		return attribute switch
		{
			"min" => eExcelConditionalFormattingValueObjectType.Min, 
			"max" => eExcelConditionalFormattingValueObjectType.Max, 
			"num" => eExcelConditionalFormattingValueObjectType.Num, 
			"formula" => eExcelConditionalFormattingValueObjectType.Formula, 
			"percent" => eExcelConditionalFormattingValueObjectType.Percent, 
			"percentile" => eExcelConditionalFormattingValueObjectType.Percentile, 
			_ => throw new Exception("Unexistent eExcelConditionalFormattingValueObjectType attribute in Conditional Formatting"), 
		};
	}

	public static XmlNode GetCfvoNodeByPosition(eExcelConditionalFormattingValueObjectPosition position, eExcelConditionalFormattingRuleType ruleType, XmlNode topNode, XmlNamespaceManager nameSpaceManager)
	{
		return topNode.SelectSingleNode(string.Format("{0}[position()={1}]", "d:cfvo", GetOrderByPosition(position, ruleType)), nameSpaceManager) ?? throw new Exception("Missing 'cfvo' node in Conditional Formatting");
	}

	public static string GetAttributeByType(eExcelConditionalFormattingValueObjectType type)
	{
		return type switch
		{
			eExcelConditionalFormattingValueObjectType.Min => "min", 
			eExcelConditionalFormattingValueObjectType.Max => "max", 
			eExcelConditionalFormattingValueObjectType.Num => "num", 
			eExcelConditionalFormattingValueObjectType.Formula => "formula", 
			eExcelConditionalFormattingValueObjectType.Percent => "percent", 
			eExcelConditionalFormattingValueObjectType.Percentile => "percentile", 
			_ => string.Empty, 
		};
	}

	public static string GetParentPathByRuleType(eExcelConditionalFormattingRuleType ruleType)
	{
		switch (ruleType)
		{
		case eExcelConditionalFormattingRuleType.ThreeColorScale:
		case eExcelConditionalFormattingRuleType.TwoColorScale:
			return "d:colorScale";
		case eExcelConditionalFormattingRuleType.ThreeIconSet:
		case eExcelConditionalFormattingRuleType.FourIconSet:
		case eExcelConditionalFormattingRuleType.FiveIconSet:
			return "d:iconSet";
		case eExcelConditionalFormattingRuleType.DataBar:
			return "d:dataBar";
		default:
			return string.Empty;
		}
	}

	public static string GetNodePathByNodeType(eExcelConditionalFormattingValueObjectNodeType nodeType)
	{
		return nodeType switch
		{
			eExcelConditionalFormattingValueObjectNodeType.Cfvo => "d:cfvo", 
			eExcelConditionalFormattingValueObjectNodeType.Color => "d:color", 
			_ => string.Empty, 
		};
	}
}
