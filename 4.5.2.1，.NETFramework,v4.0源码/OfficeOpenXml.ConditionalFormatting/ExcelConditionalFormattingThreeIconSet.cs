using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting;

public class ExcelConditionalFormattingThreeIconSet : ExcelConditionalFormattingIconSetBase<eExcelconditionalFormatting3IconsSetType>
{
	internal ExcelConditionalFormattingThreeIconSet(ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(eExcelConditionalFormattingRuleType.ThreeIconSet, address, priority, worksheet, itemElementNode, (namespaceManager == null) ? worksheet.NameSpaceManager : namespaceManager)
	{
	}
}
