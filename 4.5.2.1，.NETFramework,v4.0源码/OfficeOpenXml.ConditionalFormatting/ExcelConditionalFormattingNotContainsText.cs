using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting;

public class ExcelConditionalFormattingNotContainsText : ExcelConditionalFormattingRule, IExcelConditionalFormattingNotContainsText, IExcelConditionalFormattingRule, IExcelConditionalFormattingWithText
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
			base.Formula = string.Format("ISERROR(SEARCH(\"{1}\",{0}))", base.Address.Start.Address, value.Replace("\"", "\"\""));
		}
	}

	internal ExcelConditionalFormattingNotContainsText(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(eExcelConditionalFormattingRuleType.NotContainsText, address, priority, worksheet, itemElementNode, (namespaceManager == null) ? worksheet.NameSpaceManager : namespaceManager)
	{
		if (itemElementNode == null)
		{
			base.Operator = eExcelConditionalFormattingOperatorType.NotContains;
			Text = string.Empty;
		}
	}

	internal ExcelConditionalFormattingNotContainsText(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode)
		: this(address, priority, worksheet, itemElementNode, null)
	{
	}

	internal ExcelConditionalFormattingNotContainsText(ExcelAddress address, int priority, ExcelWorksheet worksheet)
		: this(address, priority, worksheet, null, null)
	{
	}
}
