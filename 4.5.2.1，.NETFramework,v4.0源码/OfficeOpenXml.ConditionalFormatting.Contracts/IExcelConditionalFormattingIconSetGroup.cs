namespace OfficeOpenXml.ConditionalFormatting.Contracts;

public interface IExcelConditionalFormattingIconSetGroup<T> : IExcelConditionalFormattingRule
{
	bool Reverse { get; set; }

	bool ShowValue { get; set; }

	T IconSet { get; set; }
}
