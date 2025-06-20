using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting;

public class ExcelConditionalFormattingAboveAverage : ExcelConditionalFormattingAverageGroup
{
	internal ExcelConditionalFormattingAboveAverage(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(eExcelConditionalFormattingRuleType.AboveAverage, address, priority, worksheet, itemElementNode, (namespaceManager == null) ? worksheet.NameSpaceManager : namespaceManager)
	{
		if (itemElementNode == null)
		{
			base.AboveAverage = true;
			base.EqualAverage = false;
		}
	}

	internal ExcelConditionalFormattingAboveAverage(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode)
		: this(address, priority, worksheet, itemElementNode, null)
	{
	}

	internal ExcelConditionalFormattingAboveAverage(ExcelAddress address, int priority, ExcelWorksheet worksheet)
		: this(address, priority, worksheet, null, null)
	{
	}
}
