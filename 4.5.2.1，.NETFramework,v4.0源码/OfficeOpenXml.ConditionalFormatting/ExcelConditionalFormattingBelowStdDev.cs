using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting;

public class ExcelConditionalFormattingBelowStdDev : ExcelConditionalFormattingRule, IExcelConditionalFormattingStdDevGroup, IExcelConditionalFormattingRule, IExcelConditionalFormattingWithStdDev
{
	internal ExcelConditionalFormattingBelowStdDev(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(eExcelConditionalFormattingRuleType.BelowStdDev, address, priority, worksheet, itemElementNode, (namespaceManager == null) ? worksheet.NameSpaceManager : namespaceManager)
	{
		if (itemElementNode == null)
		{
			base.AboveAverage = false;
			base.StdDev = 1;
		}
	}

	internal ExcelConditionalFormattingBelowStdDev(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode)
		: this(address, priority, worksheet, itemElementNode, null)
	{
	}

	internal ExcelConditionalFormattingBelowStdDev(ExcelAddress address, int priority, ExcelWorksheet worksheet)
		: this(address, priority, worksheet, null, null)
	{
	}
}
