namespace OfficeOpenXml.ConditionalFormatting.Contracts;

public interface IExcelConditionalFormattingThreeIconSet<T> : IExcelConditionalFormattingIconSetGroup<T>, IExcelConditionalFormattingRule
{
	ExcelConditionalFormattingIconDataBarValue Icon1 { get; }

	ExcelConditionalFormattingIconDataBarValue Icon2 { get; }

	ExcelConditionalFormattingIconDataBarValue Icon3 { get; }
}
