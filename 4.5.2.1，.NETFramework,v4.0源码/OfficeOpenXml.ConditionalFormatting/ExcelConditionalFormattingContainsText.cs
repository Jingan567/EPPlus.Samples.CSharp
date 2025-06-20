using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting;

public class ExcelConditionalFormattingContainsText : ExcelConditionalFormattingRule, IExcelConditionalFormattingContainsText, IExcelConditionalFormattingRule, IExcelConditionalFormattingWithText
{
	public string Text
	{
		get
		{
			return GetXmlNodeString("@text");
		}
		set
		{
			SetXmlNodeString("@text", value);
			base.Formula = string.Format("NOT(ISERROR(SEARCH(\"{1}\",{0})))", base.Address.Start.Address, value.Replace("\"", "\"\""));
		}
	}

	internal ExcelConditionalFormattingContainsText(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(eExcelConditionalFormattingRuleType.ContainsText, address, priority, worksheet, itemElementNode, (namespaceManager == null) ? worksheet.NameSpaceManager : namespaceManager)
	{
		if (itemElementNode == null)
		{
			base.Operator = eExcelConditionalFormattingOperatorType.ContainsText;
			Text = string.Empty;
		}
	}

	internal ExcelConditionalFormattingContainsText(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode)
		: this(address, priority, worksheet, itemElementNode, null)
	{
	}

	internal ExcelConditionalFormattingContainsText(ExcelAddress address, int priority, ExcelWorksheet worksheet)
		: this(address, priority, worksheet, null, null)
	{
	}
}
