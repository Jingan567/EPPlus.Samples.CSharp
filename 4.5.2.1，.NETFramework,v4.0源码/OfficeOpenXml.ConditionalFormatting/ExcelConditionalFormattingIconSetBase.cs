using System;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting;

public class ExcelConditionalFormattingIconSetBase<T> : ExcelConditionalFormattingRule, IExcelConditionalFormattingThreeIconSet<T>, IExcelConditionalFormattingIconSetGroup<T>, IExcelConditionalFormattingRule
{
	private const string _reversePath = "d:iconSet/@reverse";

	private const string _showValuePath = "d:iconSet/@showValue";

	private const string _iconSetPath = "d:iconSet/@iconSet";

	public ExcelConditionalFormattingIconDataBarValue Icon1 { get; internal set; }

	public ExcelConditionalFormattingIconDataBarValue Icon2 { get; internal set; }

	public ExcelConditionalFormattingIconDataBarValue Icon3 { get; internal set; }

	public bool Reverse
	{
		get
		{
			return GetXmlNodeBool("d:iconSet/@reverse", blankValue: false);
		}
		set
		{
			SetXmlNodeBool("d:iconSet/@reverse", value);
		}
	}

	public bool ShowValue
	{
		get
		{
			return GetXmlNodeBool("d:iconSet/@showValue", blankValue: true);
		}
		set
		{
			SetXmlNodeBool("d:iconSet/@showValue", value);
		}
	}

	public T IconSet
	{
		get
		{
			string xmlNodeString = GetXmlNodeString("d:iconSet/@iconSet");
			xmlNodeString = xmlNodeString.Substring(1);
			return (T)Enum.Parse(typeof(T), xmlNodeString, ignoreCase: true);
		}
		set
		{
			SetXmlNodeString("d:iconSet/@iconSet", GetIconSetString(value));
		}
	}

	internal ExcelConditionalFormattingIconSetBase(eExcelConditionalFormattingRuleType type, ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
		: base(type, address, priority, worksheet, itemElementNode, (namespaceManager == null) ? worksheet.NameSpaceManager : namespaceManager)
	{
		if (itemElementNode != null && itemElementNode.HasChildNodes)
		{
			int num = 1;
			{
				foreach (XmlNode item in itemElementNode.SelectNodes("d:iconSet/d:cfvo", base.NameSpaceManager))
				{
					switch (num)
					{
					case 1:
						Icon1 = new ExcelConditionalFormattingIconDataBarValue(type, address, worksheet, item, namespaceManager);
						break;
					case 2:
						Icon2 = new ExcelConditionalFormattingIconDataBarValue(type, address, worksheet, item, namespaceManager);
						break;
					case 3:
						Icon3 = new ExcelConditionalFormattingIconDataBarValue(type, address, worksheet, item, namespaceManager);
						break;
					default:
						return;
					}
					num++;
				}
				return;
			}
		}
		XmlNode xmlNode = CreateComplexNode(base.Node, "d:iconSet");
		double num2 = type switch
		{
			eExcelConditionalFormattingRuleType.ThreeIconSet => 3.0, 
			eExcelConditionalFormattingRuleType.FourIconSet => 4.0, 
			_ => 5.0, 
		};
		XmlElement xmlElement = xmlNode.OwnerDocument.CreateElement("d:cfvo", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
		xmlNode.AppendChild(xmlElement);
		Icon1 = new ExcelConditionalFormattingIconDataBarValue(eExcelConditionalFormattingValueObjectType.Percent, 0.0, "", eExcelConditionalFormattingRuleType.ThreeIconSet, address, priority, worksheet, xmlElement, namespaceManager);
		XmlElement xmlElement2 = xmlNode.OwnerDocument.CreateElement("d:cfvo", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
		xmlNode.AppendChild(xmlElement2);
		Icon2 = new ExcelConditionalFormattingIconDataBarValue(eExcelConditionalFormattingValueObjectType.Percent, Math.Round(100.0 / num2, 0), "", eExcelConditionalFormattingRuleType.ThreeIconSet, address, priority, worksheet, xmlElement2, namespaceManager);
		XmlElement xmlElement3 = xmlNode.OwnerDocument.CreateElement("d:cfvo", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
		xmlNode.AppendChild(xmlElement3);
		Icon3 = new ExcelConditionalFormattingIconDataBarValue(eExcelConditionalFormattingValueObjectType.Percent, Math.Round(100.0 * (2.0 / num2), 0), "", eExcelConditionalFormattingRuleType.ThreeIconSet, address, priority, worksheet, xmlElement3, namespaceManager);
		base.Type = type;
	}

	internal ExcelConditionalFormattingIconSetBase(eExcelConditionalFormattingRuleType type, ExcelAddress address, int priority, ExcelWorksheet worksheet, XmlNode itemElementNode)
		: this(type, address, priority, worksheet, itemElementNode, (XmlNamespaceManager)null)
	{
	}

	internal ExcelConditionalFormattingIconSetBase(eExcelConditionalFormattingRuleType type, ExcelAddress address, int priority, ExcelWorksheet worksheet)
		: this(type, address, priority, worksheet, (XmlNode)null, (XmlNamespaceManager)null)
	{
	}

	private string GetIconSetString(T value)
	{
		if (base.Type == eExcelConditionalFormattingRuleType.FourIconSet)
		{
			return value.ToString() switch
			{
				"Arrows" => "4Arrows", 
				"ArrowsGray" => "4ArrowsGray", 
				"Rating" => "4Rating", 
				"RedToBlack" => "4RedToBlack", 
				"TrafficLights" => "4TrafficLights", 
				_ => throw new ArgumentException("Invalid type"), 
			};
		}
		if (base.Type == eExcelConditionalFormattingRuleType.FiveIconSet)
		{
			return value.ToString() switch
			{
				"Arrows" => "5Arrows", 
				"ArrowsGray" => "5ArrowsGray", 
				"Quarters" => "5Quarters", 
				"Rating" => "5Rating", 
				_ => throw new ArgumentException("Invalid type"), 
			};
		}
		return value.ToString() switch
		{
			"Arrows" => "3Arrows", 
			"ArrowsGray" => "3ArrowsGray", 
			"Flags" => "3Flags", 
			"Signs" => "3Signs", 
			"Symbols" => "3Symbols", 
			"Symbols2" => "3Symbols2", 
			"TrafficLights1" => "3TrafficLights1", 
			"TrafficLights2" => "3TrafficLights2", 
			_ => throw new ArgumentException("Invalid type"), 
		};
	}
}
