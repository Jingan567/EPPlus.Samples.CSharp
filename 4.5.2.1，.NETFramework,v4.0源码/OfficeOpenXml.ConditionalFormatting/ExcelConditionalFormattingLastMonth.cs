using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting;

public class ExcelConditionalFormattingLastMonth : ExcelConditionalFormattingTimePeriodGroup
{
	internal ExcelConditionalFormattingLastMonth(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(eExcelConditionalFormattingRuleType.LastMonth, address, priority, worksheet, itemElementNode, (namespaceManager == null) ? worksheet.NameSpaceManager : namespaceManager)
	{
		if (itemElementNode == null)
		{
			base.TimePeriod = eExcelConditionalFormattingTimePeriodType.LastMonth;
			base.Formula = string.Format("AND(MONTH({0})=MONTH(EDATE(TODAY(),0-1)),YEAR({0})=YEAR(EDATE(TODAY(),0-1)))", base.Address.Start.Address);
		}
	}

	internal ExcelConditionalFormattingLastMonth(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode)
		: this(address, priority, worksheet, itemElementNode, null)
	{
	}

	internal ExcelConditionalFormattingLastMonth(ExcelAddress address, int priority, ExcelWorksheet worksheet)
		: this(address, priority, worksheet, null, null)
	{
	}
}
