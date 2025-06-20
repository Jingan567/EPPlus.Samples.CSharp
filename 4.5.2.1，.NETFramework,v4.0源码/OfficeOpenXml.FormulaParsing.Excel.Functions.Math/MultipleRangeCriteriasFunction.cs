using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public abstract class MultipleRangeCriteriasFunction : ExcelFunction
{
	private readonly ExpressionEvaluator _expressionEvaluator;

	protected MultipleRangeCriteriasFunction()
		: this(new ExpressionEvaluator())
	{
	}

	protected MultipleRangeCriteriasFunction(ExpressionEvaluator evaluator)
	{
		OfficeOpenXml.FormulaParsing.Utilities.Require.That(evaluator).Named("evaluator").IsNotNull();
		_expressionEvaluator = evaluator;
	}

	protected bool Evaluate(object obj, string expression)
	{
		double? num = null;
		if (IsNumeric(obj))
		{
			num = ConvertUtil.GetValueDouble(obj);
		}
		if (num.HasValue)
		{
			return _expressionEvaluator.Evaluate(num.Value, expression);
		}
		return _expressionEvaluator.Evaluate(obj, expression);
	}

	protected List<int> GetMatchIndexes(ExcelDataProvider.IRangeInfo rangeInfo, string searched)
	{
		List<int> list = new List<int>();
		int num = 0;
		for (int i = rangeInfo.Address._fromRow; i <= rangeInfo.Address._toRow; i++)
		{
			for (int j = rangeInfo.Address._fromCol; j <= rangeInfo.Address._toCol; j++)
			{
				object value = rangeInfo.GetValue(i, j);
				if (searched != null && Evaluate(value, searched))
				{
					list.Add(num);
				}
				num++;
			}
		}
		return list;
	}
}
