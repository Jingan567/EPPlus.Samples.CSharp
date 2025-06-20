using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting;

public class ExcelConditionalFormattingNextMonth : ExcelConditionalFormattingTimePeriodGroup
{
	internal ExcelConditionalFormattingNextMonth(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(eExcelConditionalFormattingRuleType.NextMonth, address, priority, worksheet, itemElementNode, (namespaceManager == null) ? worksheet.NameSpaceManager : namespaceManager)
	{
		if (itemElementNode == null)
		{
			base.TimePeriod = eExcelConditionalFormattingTimePeriodType.NextMonth;
			base.Formula = string.Format("AND(MONTH({0})=MONTH(EDATE(TODAY(),0+1)), YEAR({0})=YEAR(EDATE(TODAY(),0+1)))", base.Address.Start.Address);
		}
	}

	internal ExcelConditionalFormattingNextMonth(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode)
		: this(address, priority, worksheet, itemElementNode, null)
	{
	}

	internal ExcelConditionalFormattingNextMonth(ExcelAddress address, int priority, ExcelWorksheet worksheet)
		: this(address, priority, worksheet, null, null)
	{
	}
}
