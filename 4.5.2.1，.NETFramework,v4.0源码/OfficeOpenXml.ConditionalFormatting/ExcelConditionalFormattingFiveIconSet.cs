using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting;

public class ExcelConditionalFormattingFiveIconSet : ExcelConditionalFormattingIconSetBase<eExcelconditionalFormatting5IconsSetType>, IExcelConditionalFormattingFiveIconSet, IExcelConditionalFormattingFourIconSet<eExcelconditionalFormatting5IconsSetType>, IExcelConditionalFormattingThreeIconSet<eExcelconditionalFormatting5IconsSetType>, IExcelConditionalFormattingIconSetGroup<eExcelconditionalFormatting5IconsSetType>, IExcelConditionalFormattingRule
{
	public ExcelConditionalFormattingIconDataBarValue Icon5 { get; internal set; }

	public ExcelConditionalFormattingIconDataBarValue Icon4 { get; internal set; }

	internal ExcelConditionalFormattingFiveIconSet(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(eExcelConditionalFormattingRuleType.FiveIconSet, address, priority, worksheet, itemElementNode, (namespaceManager == null) ? worksheet.NameSpaceManager : namespaceManager)
	{
		if (itemElementNode != null && itemElementNode.HasChildNodes)
		{
			XmlNode itemElementNode2 = base.TopNode.SelectSingleNode("d:iconSet/d:cfvo[position()=4]", base.NameSpaceManager);
			Icon4 = new ExcelConditionalFormattingIconDataBarValue(eExcelConditionalFormattingRuleType.FiveIconSet, address, worksheet, itemElementNode2, namespaceManager);
			XmlNode itemElementNode3 = base.TopNode.SelectSingleNode("d:iconSet/d:cfvo[position()=5]", base.NameSpaceManager);
			Icon5 = new ExcelConditionalFormattingIconDataBarValue(eExcelConditionalFormattingRuleType.FiveIconSet, address, worksheet, itemElementNode3, namespaceManager);
		}
		else
		{
			XmlNode xmlNode = base.TopNode.SelectSingleNode("d:iconSet", base.NameSpaceManager);
			XmlElement xmlElement = xmlNode.OwnerDocument.CreateElement("d:cfvo", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
			xmlNode.AppendChild(xmlElement);
			Icon4 = new ExcelConditionalFormattingIconDataBarValue(eExcelConditionalFormattingValueObjectType.Percent, 60.0, "", eExcelConditionalFormattingRuleType.ThreeIconSet, address, priority, worksheet, xmlElement, namespaceManager);
			XmlElement xmlElement2 = xmlNode.OwnerDocument.CreateElement("d:cfvo", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
			xmlNode.AppendChild(xmlElement2);
			Icon5 = new ExcelConditionalFormattingIconDataBarValue(eExcelConditionalFormattingValueObjectType.Percent, 80.0, "", eExcelConditionalFormattingRuleType.ThreeIconSet, address, priority, worksheet, xmlElement2, namespaceManager);
		}
	}

	internal ExcelConditionalFormattingFiveIconSet(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode)
		: this(address, priority, worksheet, itemElementNode, null)
	{
	}

	internal ExcelConditionalFormattingFiveIconSet(ExcelAddress address, int priority, ExcelWorksheet worksheet)
		: this(address, priority, worksheet, null, null)
	{
	}
}
