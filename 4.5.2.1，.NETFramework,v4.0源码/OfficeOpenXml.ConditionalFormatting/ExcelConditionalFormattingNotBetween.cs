using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting;

public class ExcelConditionalFormattingNotBetween : ExcelConditionalFormattingRule, IExcelConditionalFormattingNotBetween, IExcelConditionalFormattingRule, IExcelConditionalFormattingWithFormula2, IExcelConditionalFormattingWithFormula
{
	internal ExcelConditionalFormattingNotBetween(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(eExcelConditionalFormattingRuleType.NotBetween, address, priority, worksheet, itemElementNode, (namespaceManager == null) ? worksheet.NameSpaceManager : namespaceManager)
	{
		if (itemElementNode == null)
		{
			base.Operator = eExcelConditionalFormattingOperatorType.NotBetween;
			base.Formula = string.Empty;
			base.Formula2 = string.Empty;
		}
	}

	internal ExcelConditionalFormattingNotBetween(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode)
		: this(address, priority, worksheet, itemElementNode, null)
	{
	}

	internal ExcelConditionalFormattingNotBetween(ExcelAddress address, int priority, ExcelWorksheet worksheet)
		: this(address, priority, worksheet, null, null)
	{
	}
}
