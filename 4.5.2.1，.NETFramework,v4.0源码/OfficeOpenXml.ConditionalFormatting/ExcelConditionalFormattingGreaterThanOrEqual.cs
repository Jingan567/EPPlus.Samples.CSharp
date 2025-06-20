using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting;

public class ExcelConditionalFormattingGreaterThanOrEqual : ExcelConditionalFormattingRule, IExcelConditionalFormattingGreaterThanOrEqual, IExcelConditionalFormattingRule, IExcelConditionalFormattingWithFormula
{
	internal ExcelConditionalFormattingGreaterThanOrEqual(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(eExcelConditionalFormattingRuleType.GreaterThanOrEqual, address, priority, worksheet, itemElementNode, (namespaceManager == null) ? worksheet.NameSpaceManager : namespaceManager)
	{
		if (itemElementNode == null)
		{
			base.Operator = eExcelConditionalFormattingOperatorType.GreaterThanOrEqual;
			base.Formula = string.Empty;
		}
	}

	internal ExcelConditionalFormattingGreaterThanOrEqual(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode)
		: this(address, priority, worksheet, itemElementNode, null)
	{
	}

	internal ExcelConditionalFormattingGreaterThanOrEqual(ExcelAddress address, int priority, ExcelWorksheet worksheet)
		: this(address, priority, worksheet, null, null)
	{
	}
}
