using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting;

public class ExcelConditionalFormattingEndsWith : ExcelConditionalFormattingRule, IExcelConditionalFormattingEndsWith, IExcelConditionalFormattingRule, IExcelConditionalFormattingWithText
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
			base.Formula = string.Format("RIGHT({0},LEN(\"{1}\"))=\"{1}\"", base.Address.Start.Address, value.Replace("\"", "\"\""));
		}
	}

	internal ExcelConditionalFormattingEndsWith(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(eExcelConditionalFormattingRuleType.EndsWith, address, priority, worksheet, itemElementNode, (namespaceManager == null) ? worksheet.NameSpaceManager : namespaceManager)
	{
		if (itemElementNode == null)
		{
			base.Operator = eExcelConditionalFormattingOperatorType.EndsWith;
			Text = string.Empty;
		}
	}

	internal ExcelConditionalFormattingEndsWith(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode)
		: this(address, priority, worksheet, itemElementNode, null)
	{
	}

	internal ExcelConditionalFormattingEndsWith(ExcelAddress address, int priority, ExcelWorksheet worksheet)
		: this(address, priority, worksheet, null, null)
	{
	}
}
