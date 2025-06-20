using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting;

public class ExcelConditionalFormattingFourIconSet : ExcelConditionalFormattingIconSetBase<eExcelconditionalFormatting4IconsSetType>, IExcelConditionalFormattingFourIconSet<eExcelconditionalFormatting4IconsSetType>, IExcelConditionalFormattingThreeIconSet<eExcelconditionalFormatting4IconsSetType>, IExcelConditionalFormattingIconSetGroup<eExcelconditionalFormatting4IconsSetType>, IExcelConditionalFormattingRule
{
	public ExcelConditionalFormattingIconDataBarValue Icon4 { get; internal set; }

	internal ExcelConditionalFormattingFourIconSet(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(eExcelConditionalFormattingRuleType.FourIconSet, address, priority, worksheet, itemElementNode, (namespaceManager == null) ? worksheet.NameSpaceManager : namespaceManager)
	{
		if (itemElementNode != null && itemElementNode.HasChildNodes)
		{
			XmlNode itemElementNode2 = base.TopNode.SelectSingleNode("d:iconSet/d:cfvo[position()=4]", base.NameSpaceManager);
			Icon4 = new ExcelConditionalFormattingIconDataBarValue(eExcelConditionalFormattingRuleType.FourIconSet, address, worksheet, itemElementNode2, namespaceManager);
			return;
		}
		XmlNode xmlNode = base.TopNode.SelectSingleNode("d:iconSet", base.NameSpaceManager);
		XmlElement xmlElement = xmlNode.OwnerDocument.CreateElement("d:cfvo", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
		xmlNode.AppendChild(xmlElement);
		Icon4 = new ExcelConditionalFormattingIconDataBarValue(eExcelConditionalFormattingValueObjectType.Percent, 75.0, "", eExcelConditionalFormattingRuleType.ThreeIconSet, address, priority, worksheet, xmlElement, namespaceManager);
	}

	internal ExcelConditionalFormattingFourIconSet(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode)
		: this(address, priority, worksheet, itemElementNode, null)
	{
	}

	internal ExcelConditionalFormattingFourIconSet(ExcelAddress address, int priority, ExcelWorksheet worksheet)
		: this(address, priority, worksheet, null, null)
	{
	}
}
