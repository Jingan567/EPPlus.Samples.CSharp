using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting;

public class ExcelConditionalFormattingThreeColorScale : ExcelConditionalFormattingRule, IExcelConditionalFormattingThreeColorScale, IExcelConditionalFormattingTwoColorScale, IExcelConditionalFormattingColorScaleGroup, IExcelConditionalFormattingRule
{
	private ExcelConditionalFormattingColorScaleValue _lowValue;

	private ExcelConditionalFormattingColorScaleValue _middleValue;

	private ExcelConditionalFormattingColorScaleValue _highValue;

	public ExcelConditionalFormattingColorScaleValue LowValue
	{
		get
		{
			return _lowValue;
		}
		set
		{
			_lowValue = value;
		}
	}

	public ExcelConditionalFormattingColorScaleValue MiddleValue
	{
		get
		{
			return _middleValue;
		}
		set
		{
			_middleValue = value;
		}
	}

	public ExcelConditionalFormattingColorScaleValue HighValue
	{
		get
		{
			return _highValue;
		}
		set
		{
			_highValue = value;
		}
	}

	internal ExcelConditionalFormattingThreeColorScale(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(eExcelConditionalFormattingRuleType.ThreeColorScale, address, priority, worksheet, itemElementNode, (namespaceManager == null) ? worksheet.NameSpaceManager : namespaceManager)
	{
		if (itemElementNode == null)
		{
			CreateComplexNode(base.Node, "d:colorScale");
			LowValue = new ExcelConditionalFormattingColorScaleValue(eExcelConditionalFormattingValueObjectPosition.Low, eExcelConditionalFormattingValueObjectType.Min, ExcelConditionalFormattingConstants.Colors.CfvoLowValue, eExcelConditionalFormattingRuleType.ThreeColorScale, address, priority, worksheet, base.NameSpaceManager);
			MiddleValue = new ExcelConditionalFormattingColorScaleValue(eExcelConditionalFormattingValueObjectPosition.Middle, eExcelConditionalFormattingValueObjectType.Percent, ExcelConditionalFormattingConstants.Colors.CfvoMiddleValue, 50.0, string.Empty, eExcelConditionalFormattingRuleType.ThreeColorScale, address, priority, worksheet, base.NameSpaceManager);
			HighValue = new ExcelConditionalFormattingColorScaleValue(eExcelConditionalFormattingValueObjectPosition.High, eExcelConditionalFormattingValueObjectType.Max, ExcelConditionalFormattingConstants.Colors.CfvoHighValue, eExcelConditionalFormattingRuleType.ThreeColorScale, address, priority, worksheet, base.NameSpaceManager);
		}
	}

	internal ExcelConditionalFormattingThreeColorScale(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode)
		: this(address, priority, worksheet, itemElementNode, null)
	{
	}

	internal ExcelConditionalFormattingThreeColorScale(ExcelAddress address, int priority, ExcelWorksheet worksheet)
		: this(address, priority, worksheet, null, null)
	{
	}
}
