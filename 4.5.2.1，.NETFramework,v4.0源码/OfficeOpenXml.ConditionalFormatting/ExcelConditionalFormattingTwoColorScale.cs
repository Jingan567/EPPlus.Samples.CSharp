using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting;

public class ExcelConditionalFormattingTwoColorScale : ExcelConditionalFormattingRule, IExcelConditionalFormattingTwoColorScale, IExcelConditionalFormattingColorScaleGroup, IExcelConditionalFormattingRule
{
	private ExcelConditionalFormattingColorScaleValue _lowValue;

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

	internal ExcelConditionalFormattingTwoColorScale(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(eExcelConditionalFormattingRuleType.TwoColorScale, address, priority, worksheet, itemElementNode, (namespaceManager == null) ? worksheet.NameSpaceManager : namespaceManager)
	{
		if (itemElementNode == null)
		{
			CreateComplexNode(base.Node, "d:colorScale");
			LowValue = new ExcelConditionalFormattingColorScaleValue(eExcelConditionalFormattingValueObjectPosition.Low, eExcelConditionalFormattingValueObjectType.Min, ExcelConditionalFormattingConstants.Colors.CfvoLowValue, eExcelConditionalFormattingRuleType.TwoColorScale, address, priority, worksheet, base.NameSpaceManager);
			HighValue = new ExcelConditionalFormattingColorScaleValue(eExcelConditionalFormattingValueObjectPosition.High, eExcelConditionalFormattingValueObjectType.Max, ExcelConditionalFormattingConstants.Colors.CfvoHighValue, eExcelConditionalFormattingRuleType.TwoColorScale, address, priority, worksheet, base.NameSpaceManager);
		}
	}

	internal ExcelConditionalFormattingTwoColorScale(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode)
		: this(address, priority, worksheet, itemElementNode, null)
	{
	}

	internal ExcelConditionalFormattingTwoColorScale(ExcelAddress address, int priority, ExcelWorksheet worksheet)
		: this(address, priority, worksheet, null, null)
	{
	}
}
