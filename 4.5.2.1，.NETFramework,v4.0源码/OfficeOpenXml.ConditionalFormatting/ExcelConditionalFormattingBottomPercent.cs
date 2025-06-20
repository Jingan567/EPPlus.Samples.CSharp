using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting;

public class ExcelConditionalFormattingBottomPercent : ExcelConditionalFormattingRule, IExcelConditionalFormattingTopBottomGroup, IExcelConditionalFormattingRule, IExcelConditionalFormattingWithRank
{
	internal ExcelConditionalFormattingBottomPercent(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(eExcelConditionalFormattingRuleType.BottomPercent, address, priority, worksheet, itemElementNode, (namespaceManager == null) ? worksheet.NameSpaceManager : namespaceManager)
	{
		if (itemElementNode == null)
		{
			base.Bottom = true;
			base.Percent = true;
			base.Rank = 10;
		}
	}

	internal ExcelConditionalFormattingBottomPercent(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode)
		: this(address, priority, worksheet, itemElementNode, null)
	{
	}

	internal ExcelConditionalFormattingBottomPercent(ExcelAddress address, int priority, ExcelWorksheet worksheet)
		: this(address, priority, worksheet, null, null)
	{
	}
}
