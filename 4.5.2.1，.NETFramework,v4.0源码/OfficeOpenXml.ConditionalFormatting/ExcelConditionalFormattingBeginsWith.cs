using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting;

public class ExcelConditionalFormattingBeginsWith : ExcelConditionalFormattingRule, IExcelConditionalFormattingBeginsWith, IExcelConditionalFormattingRule, IExcelConditionalFormattingWithText
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
			base.Formula = string.Format("LEFT({0},LEN(\"{1}\"))=\"{1}\"", base.Address.Start.Address, value.Replace("\"", "\"\""));
		}
	}

	internal ExcelConditionalFormattingBeginsWith(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(eExcelConditionalFormattingRuleType.BeginsWith, address, priority, worksheet, itemElementNode, (namespaceManager == null) ? worksheet.NameSpaceManager : namespaceManager)
	{
		if (itemElementNode == null)
		{
			base.Operator = eExcelConditionalFormattingOperatorType.BeginsWith;
			Text = string.Empty;
		}
	}

	internal ExcelConditionalFormattingBeginsWith(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode)
		: this(address, priority, worksheet, itemElementNode, null)
	{
	}

	internal ExcelConditionalFormattingBeginsWith(ExcelAddress address, int priority, ExcelWorksheet worksheet)
		: this(address, priority, worksheet, null, null)
	{
	}
}
