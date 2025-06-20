using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting;

public class ExcelConditionalFormattingLessThanOrEqual : ExcelConditionalFormattingRule, IExcelConditionalFormattingLessThanOrEqual, IExcelConditionalFormattingRule, IExcelConditionalFormattingWithFormula
{
	internal ExcelConditionalFormattingLessThanOrEqual(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(eExcelConditionalFormattingRuleType.LessThanOrEqual, address, priority, worksheet, itemElementNode, (namespaceManager == null) ? worksheet.NameSpaceManager : namespaceManager)
	{
		if (itemElementNode == null)
		{
			base.Operator = eExcelConditionalFormattingOperatorType.LessThanOrEqual;
			base.Formula = string.Empty;
		}
	}

	internal ExcelConditionalFormattingLessThanOrEqual(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode)
		: this(address, priority, worksheet, itemElementNode, null)
	{
	}

	internal ExcelConditionalFormattingLessThanOrEqual(ExcelAddress address, int priority, ExcelWorksheet worksheet)
		: this(address, priority, worksheet, null, null)
	{
	}
}
