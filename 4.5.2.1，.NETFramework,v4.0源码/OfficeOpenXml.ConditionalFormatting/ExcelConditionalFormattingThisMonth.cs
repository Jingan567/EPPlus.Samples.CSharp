using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting;

public class ExcelConditionalFormattingThisMonth : ExcelConditionalFormattingTimePeriodGroup
{
	internal ExcelConditionalFormattingThisMonth(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(eExcelConditionalFormattingRuleType.ThisMonth, address, priority, worksheet, itemElementNode, (namespaceManager == null) ? worksheet.NameSpaceManager : namespaceManager)
	{
		if (itemElementNode == null)
		{
			base.TimePeriod = eExcelConditionalFormattingTimePeriodType.ThisMonth;
			base.Formula = string.Format("AND(MONTH({0})=MONTH(TODAY()), YEAR({0})=YEAR(TODAY()))", base.Address.Start.Address);
		}
	}

	internal ExcelConditionalFormattingThisMonth(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode)
		: this(address, priority, worksheet, itemElementNode, null)
	{
	}

	internal ExcelConditionalFormattingThisMonth(ExcelAddress address, int priority, ExcelWorksheet worksheet)
		: this(address, priority, worksheet, null, null)
	{
	}
}
