using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting;

public class ExcelConditionalFormattingNotContainsBlanks : ExcelConditionalFormattingRule, IExcelConditionalFormattingNotContainsBlanks, IExcelConditionalFormattingRule
{
	internal ExcelConditionalFormattingNotContainsBlanks(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(eExcelConditionalFormattingRuleType.NotContainsBlanks, address, priority, worksheet, itemElementNode, (namespaceManager == null) ? worksheet.NameSpaceManager : namespaceManager)
	{
		if (itemElementNode == null)
		{
			base.Formula = $"LEN(TRIM({base.Address.Start.Address}))>0";
		}
	}

	internal ExcelConditionalFormattingNotContainsBlanks(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode)
		: this(address, priority, worksheet, itemElementNode, null)
	{
	}

	internal ExcelConditionalFormattingNotContainsBlanks(ExcelAddress address, int priority, ExcelWorksheet worksheet)
		: this(address, priority, worksheet, null, null)
	{
	}
}
