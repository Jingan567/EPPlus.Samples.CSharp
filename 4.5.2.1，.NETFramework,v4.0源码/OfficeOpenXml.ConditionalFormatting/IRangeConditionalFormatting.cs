using System.Drawing;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting;

public interface IRangeConditionalFormatting
{
	IExcelConditionalFormattingAverageGroup AddAboveAverage();

	IExcelConditionalFormattingAverageGroup AddAboveOrEqualAverage();

	IExcelConditionalFormattingAverageGroup AddBelowAverage();

	IExcelConditionalFormattingAverageGroup AddBelowOrEqualAverage();

	IExcelConditionalFormattingStdDevGroup AddAboveStdDev();

	IExcelConditionalFormattingStdDevGroup AddBelowStdDev();

	IExcelConditionalFormattingTopBottomGroup AddBottom();

	IExcelConditionalFormattingTopBottomGroup AddBottomPercent();

	IExcelConditionalFormattingTopBottomGroup AddTop();

	IExcelConditionalFormattingTopBottomGroup AddTopPercent();

	IExcelConditionalFormattingTimePeriodGroup AddLast7Days();

	IExcelConditionalFormattingTimePeriodGroup AddLastMonth();

	IExcelConditionalFormattingTimePeriodGroup AddLastWeek();

	IExcelConditionalFormattingTimePeriodGroup AddNextMonth();

	IExcelConditionalFormattingTimePeriodGroup AddNextWeek();

	IExcelConditionalFormattingTimePeriodGroup AddThisMonth();

	IExcelConditionalFormattingTimePeriodGroup AddThisWeek();

	IExcelConditionalFormattingTimePeriodGroup AddToday();

	IExcelConditionalFormattingTimePeriodGroup AddTomorrow();

	IExcelConditionalFormattingTimePeriodGroup AddYesterday();

	IExcelConditionalFormattingBeginsWith AddBeginsWith();

	IExcelConditionalFormattingBetween AddBetween();

	IExcelConditionalFormattingContainsBlanks AddContainsBlanks();

	IExcelConditionalFormattingContainsErrors AddContainsErrors();

	IExcelConditionalFormattingContainsText AddContainsText();

	IExcelConditionalFormattingDuplicateValues AddDuplicateValues();

	IExcelConditionalFormattingEndsWith AddEndsWith();

	IExcelConditionalFormattingEqual AddEqual();

	IExcelConditionalFormattingExpression AddExpression();

	IExcelConditionalFormattingGreaterThan AddGreaterThan();

	IExcelConditionalFormattingGreaterThanOrEqual AddGreaterThanOrEqual();

	IExcelConditionalFormattingLessThan AddLessThan();

	IExcelConditionalFormattingLessThanOrEqual AddLessThanOrEqual();

	IExcelConditionalFormattingNotBetween AddNotBetween();

	IExcelConditionalFormattingNotContainsBlanks AddNotContainsBlanks();

	IExcelConditionalFormattingNotContainsErrors AddNotContainsErrors();

	IExcelConditionalFormattingNotContainsText AddNotContainsText();

	IExcelConditionalFormattingNotEqual AddNotEqual();

	IExcelConditionalFormattingUniqueValues AddUniqueValues();

	IExcelConditionalFormattingThreeColorScale AddThreeColorScale();

	IExcelConditionalFormattingTwoColorScale AddTwoColorScale();

	IExcelConditionalFormattingThreeIconSet<eExcelconditionalFormatting3IconsSetType> AddThreeIconSet(eExcelconditionalFormatting3IconsSetType IconSet);

	IExcelConditionalFormattingFourIconSet<eExcelconditionalFormatting4IconsSetType> AddFourIconSet(eExcelconditionalFormatting4IconsSetType IconSet);

	IExcelConditionalFormattingFiveIconSet AddFiveIconSet(eExcelconditionalFormatting5IconsSetType IconSet);

	IExcelConditionalFormattingDataBarGroup AddDatabar(Color color);
}
