using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting;

public class ExcelConditionalFormattingYesterday : ExcelConditionalFormattingTimePeriodGroup
{
	internal ExcelConditionalFormattingYesterday(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(eExcelConditionalFormattingRuleType.Yesterday, address, priority, worksheet, itemElementNode, (namespaceManager == null) ? worksheet.NameSpaceManager : namespaceManager)
	{
		if (itemElementNode == null)
		{
			base.TimePeriod = eExcelConditionalFormattingTimePeriodType.Yesterday;
			base.Formula = $"FLOOR({base.Address.Start.Address},1)=TODAY()-1";
		}
	}

	internal ExcelConditionalFormattingYesterday(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode)
		: this(address, priority, worksheet, itemElementNode, null)
	{
	}

	internal ExcelConditionalFormattingYesterday(ExcelAddress address, int priority, ExcelWorksheet worksheet)
		: this(address, priority, worksheet, null, null)
	{
	}
}
