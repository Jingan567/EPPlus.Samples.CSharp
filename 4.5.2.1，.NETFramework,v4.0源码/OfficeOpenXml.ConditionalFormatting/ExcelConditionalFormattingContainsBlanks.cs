using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting;

public class ExcelConditionalFormattingContainsBlanks : ExcelConditionalFormattingRule, IExcelConditionalFormattingContainsBlanks, IExcelConditionalFormattingRule
{
	internal ExcelConditionalFormattingContainsBlanks(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(eExcelConditionalFormattingRuleType.ContainsBlanks, address, priority, worksheet, itemElementNode, (namespaceManager == null) ? worksheet.NameSpaceManager : namespaceManager)
	{
		if (itemElementNode == null)
		{
			base.Formula = $"LEN(TRIM({base.Address.Start.Address}))=0";
		}
	}

	internal ExcelConditionalFormattingContainsBlanks(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode)
		: this(address, priority, worksheet, itemElementNode, null)
	{
	}

	internal ExcelConditionalFormattingContainsBlanks(ExcelAddress address, int priority, ExcelWorksheet worksheet)
		: this(address, priority, worksheet, null, null)
	{
	}
}
