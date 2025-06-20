using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting;

public class ExcelConditionalFormattingAverageGroup : ExcelConditionalFormattingRule, IExcelConditionalFormattingAverageGroup, IExcelConditionalFormattingRule
{
	internal ExcelConditionalFormattingAverageGroup(eExcelConditionalFormattingRuleType type, ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(type, address, priority, worksheet, itemElementNode, (namespaceManager == null) ? worksheet.NameSpaceManager : namespaceManager)
	{
	}

	internal ExcelConditionalFormattingAverageGroup(eExcelConditionalFormattingRuleType type, ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode)
		: this(type, address, priority, worksheet, itemElementNode, null)
	{
	}

	internal ExcelConditionalFormattingAverageGroup(eExcelConditionalFormattingRuleType type, ExcelAddress address, int priority, ExcelWorksheet worksheet)
		: this(type, address, priority, worksheet, null, null)
	{
	}
}
