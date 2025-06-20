using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting;

public class ExcelConditionalFormattingBetween : ExcelConditionalFormattingRule, IExcelConditionalFormattingBetween, IExcelConditionalFormattingRule, IExcelConditionalFormattingWithFormula2, IExcelConditionalFormattingWithFormula
{
	internal ExcelConditionalFormattingBetween(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(eExcelConditionalFormattingRuleType.Between, address, priority, worksheet, itemElementNode, (namespaceManager == null) ? worksheet.NameSpaceManager : namespaceManager)
	{
		if (itemElementNode == null)
		{
			base.Operator = eExcelConditionalFormattingOperatorType.Between;
			base.Formula = string.Empty;
			base.Formula2 = string.Empty;
		}
	}

	internal ExcelConditionalFormattingBetween(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode)
		: this(address, priority, worksheet, itemElementNode, null)
	{
	}

	internal ExcelConditionalFormattingBetween(ExcelAddress address, int priority, ExcelWorksheet worksheet)
		: this(address, priority, worksheet, null, null)
	{
	}
}
