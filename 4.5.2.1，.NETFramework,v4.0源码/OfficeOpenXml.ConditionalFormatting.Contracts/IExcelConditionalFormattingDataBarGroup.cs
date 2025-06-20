using System.Drawing;

namespace OfficeOpenXml.ConditionalFormatting.Contracts;

public interface IExcelConditionalFormattingDataBarGroup : IExcelConditionalFormattingRule
{
	bool ShowValue { get; set; }

	ExcelConditionalFormattingIconDataBarValue LowValue { get; }

	ExcelConditionalFormattingIconDataBarValue HighValue { get; }

	Color Color { get; set; }
}
