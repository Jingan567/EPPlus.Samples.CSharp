using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting;

public class ExcelConditionalFormattingTimePeriodGroup : ExcelConditionalFormattingRule, IExcelConditionalFormattingTimePeriodGroup, IExcelConditionalFormattingRule
{
	internal ExcelConditionalFormattingTimePeriodGroup(eExcelConditionalFormattingRuleType type, ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(type, address, priority, worksheet, itemElementNode, (namespaceManager == null) ? worksheet.NameSpaceManager : namespaceManager)
	{
	}

	internal ExcelConditionalFormattingTimePeriodGroup(eExcelConditionalFormattingRuleType type, ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode)
		: this(type, address, priority, worksheet, itemElementNode, null)
	{
	}

	internal ExcelConditionalFormattingTimePeriodGroup(eExcelConditionalFormattingRuleType type, ExcelAddress address, int priority, ExcelWorksheet worksheet)
		: this(type, address, priority, worksheet, null, null)
	{
	}
}
