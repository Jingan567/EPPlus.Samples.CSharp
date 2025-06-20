namespace OfficeOpenXml.ConditionalFormatting.Contracts;

public interface IExcelConditionalFormattingFourIconSet<T> : IExcelConditionalFormattingThreeIconSet<T>, IExcelConditionalFormattingIconSetGroup<T>, IExcelConditionalFormattingRule
{
	ExcelConditionalFormattingIconDataBarValue Icon4 { get; }
}
