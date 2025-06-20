using System;
using System.Drawing;
using System.Xml;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.ConditionalFormatting;

public class ExcelConditionalFormattingColorScaleValue : XmlHelper
{
	private eExcelConditionalFormattingValueObjectPosition _position;

	private eExcelConditionalFormattingRuleType _ruleType;

	private ExcelWorksheet _worksheet;

	internal eExcelConditionalFormattingValueObjectPosition Position
	{
		get
		{
			return _position;
		}
		set
		{
			_position = value;
		}
	}

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
			return ExcelConditionalFormattingValueObjectType.GetTypeByAttrbiute(GetXmlNodeString(string.Format("{0}[position()={1}]/{2}", "d:cfvo", GetNodeOrder(), "@type")));
		}
		set
		{
			CreateNodeByOrdem(eExcelConditionalFormattingValueObjectNodeType.Cfvo, "@type", ExcelConditionalFormattingValueObjectType.GetAttributeByType(value));
			bool flag = false;
			eExcelConditionalFormattingValueObjectType type = Type;
			if ((uint)(type - 1) <= 1u)
			{
				flag = true;
			}
			if (flag)
			{
				string nodePathByNodeType = ExcelConditionalFormattingValueObjectType.GetNodePathByNodeType(eExcelConditionalFormattingValueObjectNodeType.Cfvo);
				int nodeOrder = GetNodeOrder();
				CreateComplexNode(base.TopNode, string.Format("{0}[position()={1}]/{2}=''", nodePathByNodeType, nodeOrder, "@val"));
			}
		}
	}

	public Color Color
	{
		get
		{
			return ExcelConditionalFormattingHelper.ConvertFromColorCode(GetXmlNodeString(string.Format("{0}[position()={1}]/{2}", "d:color", GetNodeOrder(), "@rgb")));
		}
		set
		{
			CreateNodeByOrdem(eExcelConditionalFormattingValueObjectNodeType.Color, "@rgb", value.ToArgb().ToString("x"));
		}
	}

	public double Value
	{
		get
		{
			return GetXmlNodeDouble(string.Format("{0}[position()={1}]/{2}", "d:cfvo", GetNodeOrder(), "@val"));
		}
		set
		{
			string attributeValue = string.Empty;
			if (Type == eExcelConditionalFormattingValueObjectType.Num || Type == eExcelConditionalFormattingValueObjectType.Percent || Type == eExcelConditionalFormattingValueObjectType.Percentile)
			{
				attributeValue = value.ToString();
			}
			CreateNodeByOrdem(eExcelConditionalFormattingValueObjectNodeType.Cfvo, "@val", attributeValue);
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
			return GetXmlNodeString(string.Format("{0}[position()={1}]/{2}", "d:cfvo", GetNodeOrder(), "@val"));
		}
		set
		{
			if (Type == eExcelConditionalFormattingValueObjectType.Formula)
			{
				CreateNodeByOrdem(eExcelConditionalFormattingValueObjectNodeType.Cfvo, "@val", (value == null) ? string.Empty : value.ToString());
			}
		}
	}

	internal ExcelConditionalFormattingColorScaleValue(eExcelConditionalFormattingValueObjectPosition position, eExcelConditionalFormattingValueObjectType type, Color color, double value, string formula, eExcelConditionalFormattingRuleType ruleType, ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(namespaceManager, itemElementNode)
	{
		Require.Argument(priority).IsInRange(1, int.MaxValue, "priority");
		Require.Argument(address).IsNotNull("address");
		Require.Argument(worksheet).IsNotNull("worksheet");
		_worksheet = worksheet;
		base.SchemaNodeOrder = new string[2] { "cfvo", "color" };
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
		Position = position;
		RuleType = ruleType;
		Type = type;
		Color = color;
		Value = value;
		Formula = formula;
	}

	internal ExcelConditionalFormattingColorScaleValue(eExcelConditionalFormattingValueObjectPosition position, eExcelConditionalFormattingValueObjectType type, Color color, double value, string formula, eExcelConditionalFormattingRuleType ruleType, ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNamespaceManager namespaceManager)
		: this(position, type, color, value, formula, ruleType, address, priority, worksheet, null, namespaceManager)
	{
	}

	internal ExcelConditionalFormattingColorScaleValue(eExcelConditionalFormattingValueObjectPosition position, eExcelConditionalFormattingValueObjectType type, Color color, eExcelConditionalFormattingRuleType ruleType, ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNamespaceManager namespaceManager)
		: this(position, type, color, 0.0, null, ruleType, address, priority, worksheet, null, namespaceManager)
	{
	}

	private int GetNodeOrder()
	{
		return ExcelConditionalFormattingValueObjectType.GetOrderByPosition(Position, RuleType);
	}

	private void CreateNodeByOrdem(eExcelConditionalFormattingValueObjectNodeType nodeType, string attributePath, string attributeValue)
	{
		XmlNode topNode = base.TopNode;
		string nodePathByNodeType = ExcelConditionalFormattingValueObjectType.GetNodePathByNodeType(nodeType);
		int nodeOrder = GetNodeOrder();
		eNodeInsertOrder nodeInsertOrder = eNodeInsertOrder.SchemaOrder;
		XmlNode xmlNode = null;
		if (nodeOrder > 1)
		{
			xmlNode = base.TopNode.SelectSingleNode($"{nodePathByNodeType}[position()={nodeOrder - 1}]", _worksheet.NameSpaceManager);
			if (xmlNode != null)
			{
				nodeInsertOrder = eNodeInsertOrder.After;
			}
		}
		XmlNode node = (base.TopNode = CreateComplexNode(base.TopNode, $"{nodePathByNodeType}[position()={nodeOrder}]", nodeInsertOrder, xmlNode));
		SetXmlNodeString(node, attributePath, attributeValue, removeIfBlank: true);
		base.TopNode = topNode;
	}
}
