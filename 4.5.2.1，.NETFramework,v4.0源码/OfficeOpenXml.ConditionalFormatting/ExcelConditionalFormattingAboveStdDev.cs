using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting;

public class ExcelConditionalFormattingAboveStdDev : ExcelConditionalFormattingRule, IExcelConditionalFormattingStdDevGroup, IExcelConditionalFormattingRule, IExcelConditionalFormattingWithStdDev
{
	internal ExcelConditionalFormattingAboveStdDev(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(eExcelConditionalFormattingRuleType.AboveStdDev, address, priority, worksheet, itemElementNode, (namespaceManager == null) ? worksheet.NameSpaceManager : namespaceManager)
	{
		if (itemElementNode == null)
		{
			base.AboveAverage = true;
			base.StdDev = 1;
		}
	}

	internal ExcelConditionalFormattingAboveStdDev(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode)
		: this(address, priority, worksheet, itemElementNode, null)
	{
	}

	internal ExcelConditionalFormattingAboveStdDev(ExcelAddress address, int priority, ExcelWorksheet worksheet)
		: this(address, priority, worksheet, null, null)
	{
	}
}
