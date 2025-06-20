using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting;

public class ExcelConditionalFormattingTomorrow : ExcelConditionalFormattingTimePeriodGroup
{
	internal ExcelConditionalFormattingTomorrow(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(eExcelConditionalFormattingRuleType.Tomorrow, address, priority, worksheet, itemElementNode, (namespaceManager == null) ? worksheet.NameSpaceManager : namespaceManager)
	{
		if (itemElementNode == null)
		{
			base.TimePeriod = eExcelConditionalFormattingTimePeriodType.Tomorrow;
			base.Formula = $"FLOOR({base.Address.Start.Address},1)=TODAY()+1";
		}
	}

	internal ExcelConditionalFormattingTomorrow(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode)
		: this(address, priority, worksheet, itemElementNode, null)
	{
	}

	internal ExcelConditionalFormattingTomorrow(ExcelAddress address, int priority, ExcelWorksheet worksheet)
		: this(address, priority, worksheet, null, null)
	{
	}
}
