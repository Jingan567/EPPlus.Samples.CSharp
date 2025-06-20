using System;
using System.Drawing;
using System.Globalization;
using System.Xml;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.ConditionalFormatting;

public class ExcelConditionalFormattingIconDataBarValue : XmlHelper
{
	private eExcelConditionalFormattingRuleType _ruleType;

	private ExcelWorksheet _worksheet;

	internal eExcelConditionalFormattingRuleType RuleType
	{
		get
		{
			return _ruleType;
		}
		set
		{
			_ruleType = value;
		}
	}

	public eExcelConditionalFormattingValueObjectType Type
	{
		get
		{
			return ExcelConditionalFormattingValueObjectType.GetTypeByAttrbiute(GetXmlNodeString("@type"));
		}
		set
		{
			if ((_ruleType == eExcelConditionalFormattingRuleType.ThreeIconSet || _ruleType == eExcelConditionalFormattingRuleType.FourIconSet || _ruleType == eExcelConditionalFormattingRuleType.FiveIconSet) && (value == eExcelConditionalFormattingValueObjectType.Min || value == eExcelConditionalFormattingValueObjectType.Max))
			{
				throw new ArgumentException("Value type can't be Min or Max for icon sets");
			}
			SetXmlNodeString("@type", value.ToString().ToLower(CultureInfo.InvariantCulture));
		}
	}

	public bool GreaterThanOrEqualTo
	{
		get
		{
			return GetXmlNodeBool("@gte");
		}
		set
		{
			SetXmlNodeString("@gte", (!value) ? "0" : string.Empty, removeIfBlank: true);
		}
	}

	public double Value
	{
		get
		{
			if (Type == eExcelConditionalFormattingValueObjectType.Num || Type == eExcelConditionalFormattingValueObjectType.Percent || Type == eExcelConditionalFormattingValueObjectType.Percentile)
			{
				return GetXmlNodeDouble("@val");
			}
			return 0.0;
		}
		set
		{
			string value2 = string.Empty;
			if (Type == eExcelConditionalFormattingValueObjectType.Num || Type == eExcelConditionalFormattingValueObjectType.Percent || Type == eExcelConditionalFormattingValueObjectType.Percentile)
			{
				value2 = value.ToString(CultureInfo.InvariantCulture);
			}
			SetXmlNodeString("@val", value2);
		}
	}

	public string Formula
	{
		get
		{
			if (Type != 0)
			{
				return string.Empty;
			}
			return GetXmlNodeString("@val");
		}
		set
		{
			if (Type == eExcelConditionalFormattingValueObjectType.Formula)
			{
				SetXmlNodeString("@val", value);
			}
		}
	}

	internal ExcelConditionalFormattingIconDataBarValue(eExcelConditionalFormattingValueObjectType type, double value, string formula, eExcelConditionalFormattingRuleType ruleType, ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: this(ruleType, address, worksheet, itemElementNode, namespaceManager)
	{
		Require.Argument(priority).IsInRange(1, int.MaxValue, "priority");
		if (itemElementNode == null)
		{
			string parentPathByRuleType = ExcelConditionalFormattingValueObjectType.GetParentPathByRuleType(ruleType);
			if (parentPathByRuleType == string.Empty)
			{
				throw new Exception("Missing 'cfvo' parent node in Conditional Formatting");
			}
			itemElementNode = _worksheet.WorksheetXml.SelectSingleNode(string.Format("//{0}[{1}='{2}']/{3}[{4}='{5}']/{6}", "d:conditionalFormatting", "@sqref", address.Address, "d:cfRule", "@priority", priority, parentPathByRuleType), _worksheet.NameSpaceManager);
			if (itemElementNode == null)
			{
				throw new Exception("Missing 'cfvo' parent node in Conditional Formatting");
			}
		}
		base.TopNode = itemElementNode;
		RuleType = ruleType;
		Type = type;
		Value = value;
		Formula = formula;
	}

	internal ExcelConditionalFormattingIconDataBarValue(eExcelConditionalFormattingRuleType ruleType, ExcelAddress address, ExcelWorksheet worksheet, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(namespaceManager, itemElementNode)
	{
		Require.Argument(address).IsNotNull("address");
		Require.Argument(worksheet).IsNotNull("worksheet");
		_worksheet = worksheet;
		base.SchemaNodeOrder = new string[1] { "cfvo" };
		if (itemElementNode == null && ExcelConditionalFormattingValueObjectType.GetParentPathByRuleType(ruleType) == string.Empty)
		{
			throw new Exception("Missing 'cfvo' parent node in Conditional Formatting");
		}
		RuleType = ruleType;
	}

	internal ExcelConditionalFormattingIconDataBarValue(eExcelConditionalFormattingValueObjectType type, double value, string formula, eExcelConditionalFormattingRuleType ruleType, ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNamespaceManager namespaceManager)
		: this(type, value, formula, ruleType, address, priority, worksheet, null, namespaceManager)
	{
	}

	internal ExcelConditionalFormattingIconDataBarValue(eExcelConditionalFormattingValueObjectType type, Color color, eExcelConditionalFormattingRuleType ruleType, ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNamespaceManager namespaceManager)
		: this(type, 0.0, null, ruleType, address, priority, worksheet, null, namespaceManager)
	{
	}
}
