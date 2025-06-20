using System.Drawing;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.ConditionalFormatting;

internal class RangeConditionalFormatting : IRangeConditionalFormatting
{
	public ExcelWorksheet _worksheet;

	public ExcelAddress _address;

	public RangeConditionalFormatting(ExcelWorksheet worksheet, ExcelAddress address)
	{
		Require.Argument(worksheet).IsNotNull("worksheet");
		Require.Argument(address).IsNotNull("address");
		_worksheet = worksheet;
		_address = address;
	}

	public IExcelConditionalFormattingAverageGroup AddAboveAverage()
	{
		return _worksheet.ConditionalFormatting.AddAboveAverage(_address);
	}

	public IExcelConditionalFormattingAverageGroup AddAboveOrEqualAverage()
	{
		return _worksheet.ConditionalFormatting.AddAboveOrEqualAverage(_address);
	}

	public IExcelConditionalFormattingAverageGroup AddBelowAverage()
	{
		return _worksheet.ConditionalFormatting.AddBelowAverage(_address);
	}

	public IExcelConditionalFormattingAverageGroup AddBelowOrEqualAverage()
	{
		return _worksheet.ConditionalFormatting.AddBelowOrEqualAverage(_address);
	}

	public IExcelConditionalFormattingStdDevGroup AddAboveStdDev()
	{
		return _worksheet.ConditionalFormatting.AddAboveStdDev(_address);
	}

	public IExcelConditionalFormattingStdDevGroup AddBelowStdDev()
	{
		return _worksheet.ConditionalFormatting.AddBelowStdDev(_address);
	}

	public IExcelConditionalFormattingTopBottomGroup AddBottom()
	{
		return _worksheet.ConditionalFormatting.AddBottom(_address);
	}

	public IExcelConditionalFormattingTopBottomGroup AddBottomPercent()
	{
		return _worksheet.ConditionalFormatting.AddBottomPercent(_address);
	}

	public IExcelConditionalFormattingTopBottomGroup AddTop()
	{
		return _worksheet.ConditionalFormatting.AddTop(_address);
	}

	public IExcelConditionalFormattingTopBottomGroup AddTopPercent()
	{
		return _worksheet.ConditionalFormatting.AddTopPercent(_address);
	}

	public IExcelConditionalFormattingTimePeriodGroup AddLast7Days()
	{
		return _worksheet.ConditionalFormatting.AddLast7Days(_address);
	}

	public IExcelConditionalFormattingTimePeriodGroup AddLastMonth()
	{
		return _worksheet.ConditionalFormatting.AddLastMonth(_address);
	}

	public IExcelConditionalFormattingTimePeriodGroup AddLastWeek()
	{
		return _worksheet.ConditionalFormatting.AddLastWeek(_address);
	}

	public IExcelConditionalFormattingTimePeriodGroup AddNextMonth()
	{
		return _worksheet.ConditionalFormatting.AddNextMonth(_address);
	}

	public IExcelConditionalFormattingTimePeriodGroup AddNextWeek()
	{
		return _worksheet.ConditionalFormatting.AddNextWeek(_address);
	}

	public IExcelConditionalFormattingTimePeriodGroup AddThisMonth()
	{
		return _worksheet.ConditionalFormatting.AddThisMonth(_address);
	}

	public IExcelConditionalFormattingTimePeriodGroup AddThisWeek()
	{
		return _worksheet.ConditionalFormatting.AddThisWeek(_address);
	}

	public IExcelConditionalFormattingTimePeriodGroup AddToday()
	{
		return _worksheet.ConditionalFormatting.AddToday(_address);
	}

	public IExcelConditionalFormattingTimePeriodGroup AddTomorrow()
	{
		return _worksheet.ConditionalFormatting.AddTomorrow(_address);
	}

	public IExcelConditionalFormattingTimePeriodGroup AddYesterday()
	{
		return _worksheet.ConditionalFormatting.AddYesterday(_address);
	}

	public IExcelConditionalFormattingBeginsWith AddBeginsWith()
	{
		return _worksheet.ConditionalFormatting.AddBeginsWith(_address);
	}

	public IExcelConditionalFormattingBetween AddBetween()
	{
		return _worksheet.ConditionalFormatting.AddBetween(_address);
	}

	public IExcelConditionalFormattingContainsBlanks AddContainsBlanks()
	{
		return _worksheet.ConditionalFormatting.AddContainsBlanks(_address);
	}

	public IExcelConditionalFormattingContainsErrors AddContainsErrors()
	{
		return _worksheet.ConditionalFormatting.AddContainsErrors(_address);
	}

	public IExcelConditionalFormattingContainsText AddContainsText()
	{
		return _worksheet.ConditionalFormatting.AddContainsText(_address);
	}

	public IExcelConditionalFormattingDuplicateValues AddDuplicateValues()
	{
		return _worksheet.ConditionalFormatting.AddDuplicateValues(_address);
	}

	public IExcelConditionalFormattingEndsWith AddEndsWith()
	{
		return _worksheet.ConditionalFormatting.AddEndsWith(_address);
	}

	public IExcelConditionalFormattingEqual AddEqual()
	{
		return _worksheet.ConditionalFormatting.AddEqual(_address);
	}

	public IExcelConditionalFormattingExpression AddExpression()
	{
		return _worksheet.ConditionalFormatting.AddExpression(_address);
	}

	public IExcelConditionalFormattingGreaterThan AddGreaterThan()
	{
		return _worksheet.ConditionalFormatting.AddGreaterThan(_address);
	}

	public IExcelConditionalFormattingGreaterThanOrEqual AddGreaterThanOrEqual()
	{
		return _worksheet.ConditionalFormatting.AddGreaterThanOrEqual(_address);
	}

	public IExcelConditionalFormattingLessThan AddLessThan()
	{
		return _worksheet.ConditionalFormatting.AddLessThan(_address);
	}

	public IExcelConditionalFormattingLessThanOrEqual AddLessThanOrEqual()
	{
		return _worksheet.ConditionalFormatting.AddLessThanOrEqual(_address);
	}

	public IExcelConditionalFormattingNotBetween AddNotBetween()
	{
		return _worksheet.ConditionalFormatting.AddNotBetween(_address);
	}

	public IExcelConditionalFormattingNotContainsBlanks AddNotContainsBlanks()
	{
		return _worksheet.ConditionalFormatting.AddNotContainsBlanks(_address);
	}

	public IExcelConditionalFormattingNotContainsErrors AddNotContainsErrors()
	{
		return _worksheet.ConditionalFormatting.AddNotContainsErrors(_address);
	}

	public IExcelConditionalFormattingNotContainsText AddNotContainsText()
	{
		return _worksheet.ConditionalFormatting.AddNotContainsText(_address);
	}

	public IExcelConditionalFormattingNotEqual AddNotEqual()
	{
		return _worksheet.ConditionalFormatting.AddNotEqual(_address);
	}

	public IExcelConditionalFormattingUniqueValues AddUniqueValues()
	{
		return _worksheet.ConditionalFormatting.AddUniqueValues(_address);
	}

	public IExcelConditionalFormattingThreeColorScale AddThreeColorScale()
	{
		return (IExcelConditionalFormattingThreeColorScale)_worksheet.ConditionalFormatting.AddRule(eExcelConditionalFormattingRuleType.ThreeColorScale, _address);
	}

	public IExcelConditionalFormattingTwoColorScale AddTwoColorScale()
	{
		return (IExcelConditionalFormattingTwoColorScale)_worksheet.ConditionalFormatting.AddRule(eExcelConditionalFormattingRuleType.TwoColorScale, _address);
	}

	public IExcelConditionalFormattingThreeIconSet<eExcelconditionalFormatting3IconsSetType> AddThreeIconSet(eExcelconditionalFormatting3IconsSetType IconSet)
	{
		IExcelConditionalFormattingThreeIconSet<eExcelconditionalFormatting3IconsSetType> obj = (IExcelConditionalFormattingThreeIconSet<eExcelconditionalFormatting3IconsSetType>)_worksheet.ConditionalFormatting.AddRule(eExcelConditionalFormattingRuleType.ThreeIconSet, _address);
		obj.IconSet = IconSet;
		return obj;
	}

	public IExcelConditionalFormattingFourIconSet<eExcelconditionalFormatting4IconsSetType> AddFourIconSet(eExcelconditionalFormatting4IconsSetType IconSet)
	{
		IExcelConditionalFormattingFourIconSet<eExcelconditionalFormatting4IconsSetType> obj = (IExcelConditionalFormattingFourIconSet<eExcelconditionalFormatting4IconsSetType>)_worksheet.ConditionalFormatting.AddRule(eExcelConditionalFormattingRuleType.FourIconSet, _address);
		obj.IconSet = IconSet;
		return obj;
	}

	public IExcelConditionalFormattingFiveIconSet AddFiveIconSet(eExcelconditionalFormatting5IconsSetType IconSet)
	{
		IExcelConditionalFormattingFiveIconSet obj = (IExcelConditionalFormattingFiveIconSet)_worksheet.ConditionalFormatting.AddRule(eExcelConditionalFormattingRuleType.FiveIconSet, _address);
		obj.IconSet = IconSet;
		return obj;
	}

	public IExcelConditionalFormattingDataBarGroup AddDatabar(Color Color)
	{
		IExcelConditionalFormattingDataBarGroup obj = (IExcelConditionalFormattingDataBarGroup)_worksheet.ConditionalFormatting.AddRule(eExcelConditionalFormattingRuleType.DataBar, _address);
		obj.Color = Color;
		return obj;
	}
}
