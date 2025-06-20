using System.Xml;
using OfficeOpenXml.Style.Dxf;

namespace OfficeOpenXml.ConditionalFormatting.Contracts;

public interface IExcelConditionalFormattingRule
{
	XmlNode Node { get; }

	eExcelConditionalFormattingRuleType Type { get; }

	ExcelAddress Address { get; set; }

	int Priority { get; set; }

	bool StopIfTrue { get; set; }

	ExcelDxfStyleConditionalFormatting Style { get; }
}
