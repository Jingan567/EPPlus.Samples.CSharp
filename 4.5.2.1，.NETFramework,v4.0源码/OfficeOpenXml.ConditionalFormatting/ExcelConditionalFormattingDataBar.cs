using System;
using System.Drawing;
using System.Globalization;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting;

public class ExcelConditionalFormattingDataBar : ExcelConditionalFormattingRule, IExcelConditionalFormattingDataBarGroup, IExcelConditionalFormattingRule
{
	private const string _showValuePath = "d:dataBar/@showValue";

	private const string _colorPath = "d:dataBar/d:color/@rgb";

	public bool ShowValue
	{
		get
		{
			return GetXmlNodeBool("d:dataBar/@showValue", blankValue: true);
		}
		set
		{
			SetXmlNodeBool("d:dataBar/@showValue", value);
		}
	}

	public ExcelConditionalFormattingIconDataBarValue LowValue { get; internal set; }

	public ExcelConditionalFormattingIconDataBarValue HighValue { get; internal set; }

	public Color Color
	{
		get
		{
			string xmlNodeString = GetXmlNodeString("d:dataBar/d:color/@rgb");
			if (!string.IsNullOrEmpty(xmlNodeString))
			{
				return Color.FromArgb(int.Parse(xmlNodeString, NumberStyles.HexNumber));
			}
			return Color.White;
		}
		set
		{
			SetXmlNodeString("d:dataBar/d:color/@rgb", value.ToArgb().ToString("X"));
		}
	}

	internal ExcelConditionalFormattingDataBar(eExcelConditionalFormattingRuleType type, ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(type, address, priority, worksheet, itemElementNode, (namespaceManager == null) ? worksheet.NameSpaceManager : namespaceManager)
	{
		string[] array = base.SchemaNodeOrder;
		Array.Resize(ref array, array.Length + 2);
		array[array.Length - 2] = "cfvo";
		array[array.Length - 1] = "color";
		base.SchemaNodeOrder = array;
		if (itemElementNode != null && itemElementNode.HasChildNodes)
		{
			bool flag = false;
			foreach (XmlNode item in itemElementNode.SelectNodes("d:dataBar/d:cfvo", base.NameSpaceManager))
			{
				if (!flag)
				{
					LowValue = new ExcelConditionalFormattingIconDataBarValue(type, address, worksheet, item, namespaceManager);
					flag = true;
				}
				else
				{
					HighValue = new ExcelConditionalFormattingIconDataBarValue(type, address, worksheet, item, namespaceManager);
				}
			}
		}
		else
		{
			XmlNode xmlNode = CreateComplexNode(base.Node, "d:dataBar");
			XmlElement xmlElement = xmlNode.OwnerDocument.CreateElement("d:cfvo", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
			xmlNode.AppendChild(xmlElement);
			LowValue = new ExcelConditionalFormattingIconDataBarValue(eExcelConditionalFormattingValueObjectType.Min, 0.0, "", eExcelConditionalFormattingRuleType.DataBar, address, priority, worksheet, xmlElement, namespaceManager);
			XmlElement xmlElement2 = xmlNode.OwnerDocument.CreateElement("d:cfvo", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
			xmlNode.AppendChild(xmlElement2);
			HighValue = new ExcelConditionalFormattingIconDataBarValue(eExcelConditionalFormattingValueObjectType.Max, 0.0, "", eExcelConditionalFormattingRuleType.DataBar, address, priority, worksheet, xmlElement2, namespaceManager);
		}
		base.Type = type;
	}

	internal ExcelConditionalFormattingDataBar(eExcelConditionalFormattingRuleType type, ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode)
		: this(type, address, priority, worksheet, itemElementNode, null)
	{
	}

	internal ExcelConditionalFormattingDataBar(eExcelConditionalFormattingRuleType type, ExcelAddress address, int priority, ExcelWorksheet worksheet)
		: this(type, address, priority, worksheet, null, null)
	{
	}
}
