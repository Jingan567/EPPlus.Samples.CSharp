namespace OfficeOpenXml.ConditionalFormatting.Contracts;

public interface IExcelConditionalFormattingTwoColorScale : IExcelConditionalFormattingColorScaleGroup, IExcelConditionalFormattingRule
{
	ExcelConditionalFormattingColorScaleValue LowValue { get; set; }

	ExcelConditionalFormattingColorScaleValue HighValue { get; set; }
}
