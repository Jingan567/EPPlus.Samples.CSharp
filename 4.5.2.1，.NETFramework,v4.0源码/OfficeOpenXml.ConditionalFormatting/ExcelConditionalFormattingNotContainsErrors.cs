using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting;

public class ExcelConditionalFormattingNotContainsErrors : ExcelConditionalFormattingRule, IExcelConditionalFormattingNotContainsErrors, IExcelConditionalFormattingRule
{
	internal ExcelConditionalFormattingNotContainsErrors(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(eExcelConditionalFormattingRuleType.NotContainsErrors, address, priority, worksheet, itemElementNode, (namespaceManager == null) ? worksheet.NameSpaceManager : namespaceManager)
	{
		if (itemElementNode == null)
		{
			base.Formula = $"NOT(ISERROR({base.Address.Start.Address}))";
		}
	}

	internal ExcelConditionalFormattingNotContainsErrors(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode)
		: this(address, priority, worksheet, itemElementNode, null)
	{
	}

	internal ExcelConditionalFormattingNotContainsErrors(ExcelAddress address, int priority, ExcelWorksheet worksheet)
		: this(address, priority, worksheet, null, null)
	{
	}
}
