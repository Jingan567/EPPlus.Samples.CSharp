using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.ConditionalFormatting;

public class ExcelConditionalFormattingCollection : XmlHelper, IEnumerable<IExcelConditionalFormattingRule>, IEnumerable
{
	private List<IExcelConditionalFormattingRule> _rules = new List<IExcelConditionalFormattingRule>();

	private ExcelWorksheet _worksheet;

	public int Count => _rules.Count;

	public IExcelConditionalFormattingRule this[int index]
	{
		get
		{
			return _rules[index];
		}
		set
		{
			_rules[index] = value;
		}
	}

	internal ExcelConditionalFormattingCollection(ExcelWorksheet worksheet)
		: base(worksheet.NameSpaceManager, worksheet.WorksheetXml.DocumentElement)
	{
		Require.Argument(worksheet).IsNotNull("worksheet");
		_worksheet = worksheet;
		base.SchemaNodeOrder = _worksheet.SchemaNodeOrder;
		XmlNodeList xmlNodeList = base.TopNode.SelectNodes("//d:conditionalFormatting", _worksheet.NameSpaceManager);
		if (xmlNodeList == null || xmlNodeList.Count <= 0)
		{
			return;
		}
		foreach (XmlNode item in xmlNodeList)
		{
			if (item.Attributes["sqref"] == null)
			{
				throw new Exception("Missing 'sqref' attribute in Conditional Formatting");
			}
			ExcelAddress address = new ExcelAddress(item.Attributes["sqref"].Value);
			foreach (XmlNode item2 in item.SelectNodes("d:cfRule", _worksheet.NameSpaceManager))
			{
				if (item2.Attributes["type"] == null)
				{
					throw new Exception("Missing 'type' attribute in Conditional Formatting Rule");
				}
				if (item2.Attributes["priority"] == null)
				{
					throw new Exception("Missing 'priority' attribute in Conditional Formatting Rule");
				}
				string attributeString = ExcelConditionalFormattingHelper.GetAttributeString(item2, "type");
				ExcelConditionalFormattingRule excelConditionalFormattingRule = ExcelConditionalFormattingRuleFactory.Create(priority: ExcelConditionalFormattingHelper.GetAttributeInt(item2, "priority"), type: ExcelConditionalFormattingRuleType.GetTypeByAttrbiute(attributeString, item2, _worksheet.NameSpaceManager), address: address, worksheet: _worksheet, itemElementNode: item2);
				if (excelConditionalFormattingRule != null)
				{
					_rules.Add(excelConditionalFormattingRule);
				}
			}
		}
	}

	private void EnsureRootElementExists()
	{
		if (_worksheet.WorksheetXml.DocumentElement == null)
		{
			throw new Exception("Missing 'worksheet' node");
		}
	}

	private XmlNode GetRootNode()
	{
		EnsureRootElementExists();
		return _worksheet.WorksheetXml.DocumentElement;
	}

	private ExcelAddress ValidateAddress(ExcelAddress address)
	{
		Require.Argument(address).IsNotNull("address");
		return address;
	}

	private int GetNextPriority()
	{
		int num = 0;
		foreach (IExcelConditionalFormattingRule rule in _rules)
		{
			if (rule.Priority > num)
			{
				num = rule.Priority;
			}
		}
		return num + 1;
	}

	IEnumerator<IExcelConditionalFormattingRule> IEnumerable<IExcelConditionalFormattingRule>.GetEnumerator()
	{
		return _rules.GetEnumerator();
	}

	IEnumerator IEnumerable.GetEnumerator()
	{
		return _rules.GetEnumerator();
	}

	public void RemoveAll()
	{
		foreach (XmlNode item in base.TopNode.SelectNodes("//d:conditionalFormatting", _worksheet.NameSpaceManager))
		{
			item.ParentNode.RemoveChild(item);
		}
		_rules.Clear();
	}

	public void Remove(IExcelConditionalFormattingRule item)
	{
		Require.Argument(item).IsNotNull("item");
		try
		{
			XmlNode parentNode = item.Node.ParentNode;
			parentNode.RemoveChild(item.Node);
			if (!parentNode.HasChildNodes)
			{
				parentNode.ParentNode.RemoveChild(parentNode);
			}
			_rules.Remove(item);
		}
		catch
		{
			throw new Exception("Invalid remove rule operation");
		}
	}

	public void RemoveAt(int index)
	{
		Require.Argument(index).IsInRange(0, Count - 1, "index");
		Remove(this[index]);
	}

	public void RemoveByPriority(int priority)
	{
		try
		{
			Remove(RulesByPriority(priority));
		}
		catch
		{
		}
	}

	public IExcelConditionalFormattingRule RulesByPriority(int priority)
	{
		return _rules.Find((IExcelConditionalFormattingRule x) => x.Priority == priority);
	}

	internal IExcelConditionalFormattingRule AddRule(eExcelConditionalFormattingRuleType type, ExcelAddress address)
	{
		Require.Argument(address).IsNotNull("address");
		address = ValidateAddress(address);
		EnsureRootElementExists();
		IExcelConditionalFormattingRule excelConditionalFormattingRule = ExcelConditionalFormattingRuleFactory.Create(type, address, GetNextPriority(), _worksheet, null);
		_rules.Add(excelConditionalFormattingRule);
		return excelConditionalFormattingRule;
	}

	public IExcelConditionalFormattingAverageGroup AddAboveAverage(ExcelAddress address)
	{
		return (IExcelConditionalFormattingAverageGroup)AddRule(eExcelConditionalFormattingRuleType.AboveAverage, address);
	}

	public IExcelConditionalFormattingAverageGroup AddAboveOrEqualAverage(ExcelAddress address)
	{
		return (IExcelConditionalFormattingAverageGroup)AddRule(eExcelConditionalFormattingRuleType.AboveOrEqualAverage, address);
	}

	public IExcelConditionalFormattingAverageGroup AddBelowAverage(ExcelAddress address)
	{
		return (IExcelConditionalFormattingAverageGroup)AddRule(eExcelConditionalFormattingRuleType.BelowAverage, address);
	}

	public IExcelConditionalFormattingAverageGroup AddBelowOrEqualAverage(ExcelAddress address)
	{
		return (IExcelConditionalFormattingAverageGroup)AddRule(eExcelConditionalFormattingRuleType.BelowOrEqualAverage, address);
	}

	public IExcelConditionalFormattingStdDevGroup AddAboveStdDev(ExcelAddress address)
	{
		return (IExcelConditionalFormattingStdDevGroup)AddRule(eExcelConditionalFormattingRuleType.AboveStdDev, address);
	}

	public IExcelConditionalFormattingStdDevGroup AddBelowStdDev(ExcelAddress address)
	{
		return (IExcelConditionalFormattingStdDevGroup)AddRule(eExcelConditionalFormattingRuleType.BelowStdDev, address);
	}

	public IExcelConditionalFormattingTopBottomGroup AddBottom(ExcelAddress address)
	{
		return (IExcelConditionalFormattingTopBottomGroup)AddRule(eExcelConditionalFormattingRuleType.Bottom, address);
	}

	public IExcelConditionalFormattingTopBottomGroup AddBottomPercent(ExcelAddress address)
	{
		return (IExcelConditionalFormattingTopBottomGroup)AddRule(eExcelConditionalFormattingRuleType.BottomPercent, address);
	}

	public IExcelConditionalFormattingTopBottomGroup AddTop(ExcelAddress address)
	{
		return (IExcelConditionalFormattingTopBottomGroup)AddRule(eExcelConditionalFormattingRuleType.Top, address);
	}

	public IExcelConditionalFormattingTopBottomGroup AddTopPercent(ExcelAddress address)
	{
		return (IExcelConditionalFormattingTopBottomGroup)AddRule(eExcelConditionalFormattingRuleType.TopPercent, address);
	}

	public IExcelConditionalFormattingTimePeriodGroup AddLast7Days(ExcelAddress address)
	{
		return (IExcelConditionalFormattingTimePeriodGroup)AddRule(eExcelConditionalFormattingRuleType.Last7Days, address);
	}

	public IExcelConditionalFormattingTimePeriodGroup AddLastMonth(ExcelAddress address)
	{
		return (IExcelConditionalFormattingTimePeriodGroup)AddRule(eExcelConditionalFormattingRuleType.LastMonth, address);
	}

	public IExcelConditionalFormattingTimePeriodGroup AddLastWeek(ExcelAddress address)
	{
		return (IExcelConditionalFormattingTimePeriodGroup)AddRule(eExcelConditionalFormattingRuleType.LastWeek, address);
	}

	public IExcelConditionalFormattingTimePeriodGroup AddNextMonth(ExcelAddress address)
	{
		return (IExcelConditionalFormattingTimePeriodGroup)AddRule(eExcelConditionalFormattingRuleType.NextMonth, address);
	}

	public IExcelConditionalFormattingTimePeriodGroup AddNextWeek(ExcelAddress address)
	{
		return (IExcelConditionalFormattingTimePeriodGroup)AddRule(eExcelConditionalFormattingRuleType.NextWeek, address);
	}

	public IExcelConditionalFormattingTimePeriodGroup AddThisMonth(ExcelAddress address)
	{
		return (IExcelConditionalFormattingTimePeriodGroup)AddRule(eExcelConditionalFormattingRuleType.ThisMonth, address);
	}

	public IExcelConditionalFormattingTimePeriodGroup AddThisWeek(ExcelAddress address)
	{
		return (IExcelConditionalFormattingTimePeriodGroup)AddRule(eExcelConditionalFormattingRuleType.ThisWeek, address);
	}

	public IExcelConditionalFormattingTimePeriodGroup AddToday(ExcelAddress address)
	{
		return (IExcelConditionalFormattingTimePeriodGroup)AddRule(eExcelConditionalFormattingRuleType.Today, address);
	}

	public IExcelConditionalFormattingTimePeriodGroup AddTomorrow(ExcelAddress address)
	{
		return (IExcelConditionalFormattingTimePeriodGroup)AddRule(eExcelConditionalFormattingRuleType.Tomorrow, address);
	}

	public IExcelConditionalFormattingTimePeriodGroup AddYesterday(ExcelAddress address)
	{
		return (IExcelConditionalFormattingTimePeriodGroup)AddRule(eExcelConditionalFormattingRuleType.Yesterday, address);
	}

	public IExcelConditionalFormattingBeginsWith AddBeginsWith(ExcelAddress address)
	{
		return (IExcelConditionalFormattingBeginsWith)AddRule(eExcelConditionalFormattingRuleType.BeginsWith, address);
	}

	public IExcelConditionalFormattingBetween AddBetween(ExcelAddress address)
	{
		return (IExcelConditionalFormattingBetween)AddRule(eExcelConditionalFormattingRuleType.Between, address);
	}

	public IExcelConditionalFormattingContainsBlanks AddContainsBlanks(ExcelAddress address)
	{
		return (IExcelConditionalFormattingContainsBlanks)AddRule(eExcelConditionalFormattingRuleType.ContainsBlanks, address);
	}

	public IExcelConditionalFormattingContainsErrors AddContainsErrors(ExcelAddress address)
	{
		return (IExcelConditionalFormattingContainsErrors)AddRule(eExcelConditionalFormattingRuleType.ContainsErrors, address);
	}

	public IExcelConditionalFormattingContainsText AddContainsText(ExcelAddress address)
	{
		return (IExcelConditionalFormattingContainsText)AddRule(eExcelConditionalFormattingRuleType.ContainsText, address);
	}

	public IExcelConditionalFormattingDuplicateValues AddDuplicateValues(ExcelAddress address)
	{
		return (IExcelConditionalFormattingDuplicateValues)AddRule(eExcelConditionalFormattingRuleType.DuplicateValues, address);
	}

	public IExcelConditionalFormattingEndsWith AddEndsWith(ExcelAddress address)
	{
		return (IExcelConditionalFormattingEndsWith)AddRule(eExcelConditionalFormattingRuleType.EndsWith, address);
	}

	public IExcelConditionalFormattingEqual AddEqual(ExcelAddress address)
	{
		return (IExcelConditionalFormattingEqual)AddRule(eExcelConditionalFormattingRuleType.Equal, address);
	}

	public IExcelConditionalFormattingExpression AddExpression(ExcelAddress address)
	{
		return (IExcelConditionalFormattingExpression)AddRule(eExcelConditionalFormattingRuleType.Expression, address);
	}

	public IExcelConditionalFormattingGreaterThan AddGreaterThan(ExcelAddress address)
	{
		return (IExcelConditionalFormattingGreaterThan)AddRule(eExcelConditionalFormattingRuleType.GreaterThan, address);
	}

	public IExcelConditionalFormattingGreaterThanOrEqual AddGreaterThanOrEqual(ExcelAddress address)
	{
		return (IExcelConditionalFormattingGreaterThanOrEqual)AddRule(eExcelConditionalFormattingRuleType.GreaterThanOrEqual, address);
	}

	public IExcelConditionalFormattingLessThan AddLessThan(ExcelAddress address)
	{
		return (IExcelConditionalFormattingLessThan)AddRule(eExcelConditionalFormattingRuleType.LessThan, address);
	}

	public IExcelConditionalFormattingLessThanOrEqual AddLessThanOrEqual(ExcelAddress address)
	{
		return (IExcelConditionalFormattingLessThanOrEqual)AddRule(eExcelConditionalFormattingRuleType.LessThanOrEqual, address);
	}

	public IExcelConditionalFormattingNotBetween AddNotBetween(ExcelAddress address)
	{
		return (IExcelConditionalFormattingNotBetween)AddRule(eExcelConditionalFormattingRuleType.NotBetween, address);
	}

	public IExcelConditionalFormattingNotContainsBlanks AddNotContainsBlanks(ExcelAddress address)
	{
		return (IExcelConditionalFormattingNotContainsBlanks)AddRule(eExcelConditionalFormattingRuleType.NotContainsBlanks, address);
	}

	public IExcelConditionalFormattingNotContainsErrors AddNotContainsErrors(ExcelAddress address)
	{
		return (IExcelConditionalFormattingNotContainsErrors)AddRule(eExcelConditionalFormattingRuleType.NotContainsErrors, address);
	}

	public IExcelConditionalFormattingNotContainsText AddNotContainsText(ExcelAddress address)
	{
		return (IExcelConditionalFormattingNotContainsText)AddRule(eExcelConditionalFormattingRuleType.NotContainsText, address);
	}

	public IExcelConditionalFormattingNotEqual AddNotEqual(ExcelAddress address)
	{
		return (IExcelConditionalFormattingNotEqual)AddRule(eExcelConditionalFormattingRuleType.NotEqual, address);
	}

	public IExcelConditionalFormattingUniqueValues AddUniqueValues(ExcelAddress address)
	{
		return (IExcelConditionalFormattingUniqueValues)AddRule(eExcelConditionalFormattingRuleType.UniqueValues, address);
	}

	public IExcelConditionalFormattingThreeColorScale AddThreeColorScale(ExcelAddress address)
	{
		return (IExcelConditionalFormattingThreeColorScale)AddRule(eExcelConditionalFormattingRuleType.ThreeColorScale, address);
	}

	public IExcelConditionalFormattingTwoColorScale AddTwoColorScale(ExcelAddress address)
	{
		return (IExcelConditionalFormattingTwoColorScale)AddRule(eExcelConditionalFormattingRuleType.TwoColorScale, address);
	}

	public IExcelConditionalFormattingThreeIconSet<eExcelconditionalFormatting3IconsSetType> AddThreeIconSet(ExcelAddress Address, eExcelconditionalFormatting3IconsSetType IconSet)
	{
		IExcelConditionalFormattingThreeIconSet<eExcelconditionalFormatting3IconsSetType> obj = (IExcelConditionalFormattingThreeIconSet<eExcelconditionalFormatting3IconsSetType>)AddRule(eExcelConditionalFormattingRuleType.ThreeIconSet, Address);
		obj.IconSet = IconSet;
		return obj;
	}

	public IExcelConditionalFormattingFourIconSet<eExcelconditionalFormatting4IconsSetType> AddFourIconSet(ExcelAddress Address, eExcelconditionalFormatting4IconsSetType IconSet)
	{
		IExcelConditionalFormattingFourIconSet<eExcelconditionalFormatting4IconsSetType> obj = (IExcelConditionalFormattingFourIconSet<eExcelconditionalFormatting4IconsSetType>)AddRule(eExcelConditionalFormattingRuleType.FourIconSet, Address);
		obj.IconSet = IconSet;
		return obj;
	}

	public IExcelConditionalFormattingFiveIconSet AddFiveIconSet(ExcelAddress Address, eExcelconditionalFormatting5IconsSetType IconSet)
	{
		IExcelConditionalFormattingFiveIconSet obj = (IExcelConditionalFormattingFiveIconSet)AddRule(eExcelConditionalFormattingRuleType.FiveIconSet, Address);
		obj.IconSet = IconSet;
		return obj;
	}

	public IExcelConditionalFormattingDataBarGroup AddDatabar(ExcelAddress Address, Color color)
	{
		IExcelConditionalFormattingDataBarGroup obj = (IExcelConditionalFormattingDataBarGroup)AddRule(eExcelConditionalFormattingRuleType.DataBar, Address);
		obj.Color = color;
		return obj;
	}
}
