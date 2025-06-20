using System;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Database;

public class RowMatcher
{
	private readonly WildCardValueMatcher _wildCardValueMatcher;

	private readonly ExpressionEvaluator _expressionEvaluator;

	public RowMatcher()
		: this(new WildCardValueMatcher(), new ExpressionEvaluator())
	{
	}

	public RowMatcher(WildCardValueMatcher wildCardValueMatcher, ExpressionEvaluator expressionEvaluator)
	{
		_wildCardValueMatcher = wildCardValueMatcher;
		_expressionEvaluator = expressionEvaluator;
	}

	public bool IsMatch(ExcelDatabaseRow row, ExcelDatabaseCriteria criteria)
	{
		bool result = true;
		foreach (KeyValuePair<ExcelDatabaseCriteriaField, object> item in criteria.Items)
		{
			object obj = (item.Key.FieldIndex.HasValue ? row[item.Key.FieldIndex.Value] : row[item.Key.FieldName]);
			object value = item.Value;
			if (obj.IsNumeric() && value.IsNumeric())
			{
				if (System.Math.Abs(ConvertUtil.GetValueDouble(obj) - ConvertUtil.GetValueDouble(value)) > double.Epsilon)
				{
					return false;
				}
				continue;
			}
			string expression = value.ToString();
			if (!Evaluate(obj, expression))
			{
				return false;
			}
		}
		return result;
	}

	private bool Evaluate(object obj, string expression)
	{
		if (obj == null)
		{
			return false;
		}
		double? num = null;
		if (ConvertUtil.IsNumeric(obj))
		{
			num = ConvertUtil.GetValueDouble(obj);
		}
		if (num.HasValue)
		{
			return _expressionEvaluator.Evaluate(num.Value, expression);
		}
		return _wildCardValueMatcher.IsMatch(expression, obj.ToString()) == 0;
	}
}
