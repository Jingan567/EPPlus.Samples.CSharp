using System;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.ConditionalFormatting;

public abstract class ExcelConditionalFormattingRule : XmlHelper, IExcelConditionalFormattingRule
{
	private eExcelConditionalFormattingRuleType? _type;

	private ExcelWorksheet _worksheet;

	private static bool _changingPriority;

	internal ExcelDxfStyleConditionalFormatting _style;

	public XmlNode Node => base.TopNode;

	public ExcelAddress Address
	{
		get
		{
			return new ExcelAddress(Node.ParentNode.Attributes["sqref"].Value);
		}
		set
		{
			if (Address.Address != value.Address)
			{
				_ = Node;
				XmlNode parentNode = Node.ParentNode;
				XmlNode xmlNode = CreateComplexNode(_worksheet.WorksheetXml.DocumentElement, string.Format("{0}[{1}='{2}']/{1}='{2}'", "d:conditionalFormatting", "@sqref", value.AddressSpaceSeparated));
				base.TopNode = xmlNode.AppendChild(Node);
				if (!parentNode.HasChildNodes)
				{
					parentNode.ParentNode.RemoveChild(parentNode);
				}
			}
		}
	}

	public eExcelConditionalFormattingRuleType Type
	{
		get
		{
			if (!_type.HasValue)
			{
				_type = ExcelConditionalFormattingRuleType.GetTypeByAttrbiute(GetXmlNodeString("@type"), base.TopNode, _worksheet.NameSpaceManager);
			}
			return _type.Value;
		}
		internal set
		{
			_type = value;
			SetXmlNodeString("@type", ExcelConditionalFormattingRuleType.GetAttributeByType(value), removeIfBlank: true);
		}
	}

	public int Priority
	{
		get
		{
			return GetXmlNodeInt("@priority");
		}
		set
		{
			int priority = Priority;
			if (priority == value)
			{
				return;
			}
			if (!_changingPriority)
			{
				if (value < 1)
				{
					throw new IndexOutOfRangeException("Invalid priority number. Must be bigger than zero");
				}
				_changingPriority = true;
				if (priority > value)
				{
					for (int num = priority - 1; num >= value; num--)
					{
						IExcelConditionalFormattingRule excelConditionalFormattingRule = _worksheet.ConditionalFormatting.RulesByPriority(num);
						if (excelConditionalFormattingRule != null)
						{
							excelConditionalFormattingRule.Priority++;
						}
					}
				}
				else
				{
					for (int i = priority + 1; i <= value; i++)
					{
						IExcelConditionalFormattingRule excelConditionalFormattingRule2 = _worksheet.ConditionalFormatting.RulesByPriority(i);
						if (excelConditionalFormattingRule2 != null)
						{
							excelConditionalFormattingRule2.Priority--;
						}
					}
				}
				_changingPriority = false;
			}
			SetXmlNodeString("@priority", value.ToString(), removeIfBlank: true);
		}
	}

	public bool StopIfTrue
	{
		get
		{
			return GetXmlNodeBool("@stopIfTrue");
		}
		set
		{
			SetXmlNodeString("@stopIfTrue", value ? "1" : string.Empty, removeIfBlank: true);
		}
	}

	internal int DxfId
	{
		get
		{
			return GetXmlNodeInt("@dxfId");
		}
		set
		{
			SetXmlNodeString("@dxfId", (value == int.MinValue) ? string.Empty : value.ToString(), removeIfBlank: true);
		}
	}

	public ExcelDxfStyleConditionalFormatting Style
	{
		get
		{
			if (_style == null)
			{
				_style = new ExcelDxfStyleConditionalFormatting(base.NameSpaceManager, null, _worksheet.Workbook.Styles);
			}
			return _style;
		}
	}

	public ushort StdDev
	{
		get
		{
			return Convert.ToUInt16(GetXmlNodeInt("@stdDev"));
		}
		set
		{
			SetXmlNodeString("@stdDev", (value == 0) ? "1" : value.ToString(), removeIfBlank: true);
		}
	}

	public ushort Rank
	{
		get
		{
			return Convert.ToUInt16(GetXmlNodeInt("@rank"));
		}
		set
		{
			SetXmlNodeString("@rank", (value == 0) ? "1" : value.ToString(), removeIfBlank: true);
		}
	}

	protected internal bool? AboveAverage
	{
		get
		{
			bool? xmlNodeBoolNullable = GetXmlNodeBoolNullable("@aboveAverage");
			return xmlNodeBoolNullable == true || !xmlNodeBoolNullable.HasValue;
		}
		set
		{
			string value2 = string.Empty;
			if (_type == eExcelConditionalFormattingRuleType.BelowAverage || _type == eExcelConditionalFormattingRuleType.BelowOrEqualAverage || _type == eExcelConditionalFormattingRuleType.BelowStdDev)
			{
				value2 = "0";
			}
			SetXmlNodeString("@aboveAverage", value2, removeIfBlank: true);
		}
	}

	protected internal bool? EqualAverage
	{
		get
		{
			return GetXmlNodeBoolNullable("@equalAverage") == true;
		}
		set
		{
			string value2 = string.Empty;
			if (_type == eExcelConditionalFormattingRuleType.AboveOrEqualAverage || _type == eExcelConditionalFormattingRuleType.BelowOrEqualAverage)
			{
				value2 = "1";
			}
			SetXmlNodeString("@equalAverage", value2, removeIfBlank: true);
		}
	}

	protected internal bool? Bottom
	{
		get
		{
			return GetXmlNodeBoolNullable("@bottom") == true;
		}
		set
		{
			string value2 = string.Empty;
			if (_type == eExcelConditionalFormattingRuleType.Bottom || _type == eExcelConditionalFormattingRuleType.BottomPercent)
			{
				value2 = "1";
			}
			SetXmlNodeString("@bottom", value2, removeIfBlank: true);
		}
	}

	protected internal bool? Percent
	{
		get
		{
			return GetXmlNodeBoolNullable("@percent") == true;
		}
		set
		{
			string value2 = string.Empty;
			if (_type == eExcelConditionalFormattingRuleType.BottomPercent || _type == eExcelConditionalFormattingRuleType.TopPercent)
			{
				value2 = "1";
			}
			SetXmlNodeString("@percent", value2, removeIfBlank: true);
		}
	}

	protected internal eExcelConditionalFormattingTimePeriodType TimePeriod
	{
		get
		{
			return ExcelConditionalFormattingTimePeriodType.GetTypeByAttribute(GetXmlNodeString("@timePeriod"));
		}
		set
		{
			SetXmlNodeString("@timePeriod", ExcelConditionalFormattingTimePeriodType.GetAttributeByType(value), removeIfBlank: true);
		}
	}

	protected internal eExcelConditionalFormattingOperatorType Operator
	{
		get
		{
			return ExcelConditionalFormattingOperatorType.GetTypeByAttribute(GetXmlNodeString("@operator"));
		}
		set
		{
			SetXmlNodeString("@operator", ExcelConditionalFormattingOperatorType.GetAttributeByType(value), removeIfBlank: true);
		}
	}

	public string Formula
	{
		get
		{
			return GetXmlNodeString("d:formula");
		}
		set
		{
			SetXmlNodeString("d:formula", value);
		}
	}

	public string Formula2
	{
		get
		{
			return GetXmlNodeString(string.Format("{0}[position()=2]", "d:formula"));
		}
		set
		{
			CreateComplexNode(base.TopNode, string.Format("{0}[position()=1]", "d:formula"));
			CreateComplexNode(base.TopNode, string.Format("{0}[position()=2]", "d:formula")).InnerText = value;
		}
	}

	internal ExcelConditionalFormattingRule(eExcelConditionalFormattingRuleType type, ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(namespaceManager, itemElementNode)
	{
		Require.Argument(address).IsNotNull("address");
		Require.Argument(priority).IsInRange(0, int.MaxValue, "priority");
		Require.Argument(worksheet).IsNotNull("worksheet");
		_type = type;
		_worksheet = worksheet;
		base.SchemaNodeOrder = _worksheet.SchemaNodeOrder;
		if (itemElementNode == null)
		{
			itemElementNode = CreateComplexNode(_worksheet.WorksheetXml.DocumentElement, string.Format("{0}[{1}='{2}']/{1}='{2}'/{3}[{4}='{5}']/{4}='{5}'", "d:conditionalFormatting", "@sqref", address.AddressSpaceSeparated, "d:cfRule", "@priority", priority));
		}
		base.TopNode = itemElementNode;
		Address = address;
		Priority = priority;
		Type = type;
		if (DxfId >= 0)
		{
			worksheet.Workbook.Styles.Dxfs[DxfId].AllowChange = true;
			_style = worksheet.Workbook.Styles.Dxfs[DxfId].Clone();
		}
	}

	internal ExcelConditionalFormattingRule(eExcelConditionalFormattingRuleType type, ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNamespaceManager namespaceManager)
		: this(type, address, priority, worksheet, null, namespaceManager)
	{
	}
}
