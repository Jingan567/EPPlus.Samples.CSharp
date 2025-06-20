using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting;

public class ExcelConditionalFormattingTopPercent : ExcelConditionalFormattingRule, IExcelConditionalFormattingTopBottomGroup, IExcelConditionalFormattingRule, IExcelConditionalFormattingWithRank
{
	internal ExcelConditionalFormattingTopPercent(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(eExcelConditionalFormattingRuleType.TopPercent, address, priority, worksheet, itemElementNode, (namespaceManager == null) ? worksheet.NameSpaceManager : namespaceManager)
	{
		if (itemElementNode == null)
		{
			base.Bottom = false;
			base.Percent = true;
			base.Rank = 10;
		}
	}

	internal ExcelConditionalFormattingTopPercent(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode)
		: this(address, priority, worksheet, itemElementNode, null)
	{
	}

	internal ExcelConditionalFormattingTopPercent(ExcelAddress address, int priority, ExcelWorksheet worksheet)
		: this(address, priority, worksheet, null, null)
	{
	}
}
