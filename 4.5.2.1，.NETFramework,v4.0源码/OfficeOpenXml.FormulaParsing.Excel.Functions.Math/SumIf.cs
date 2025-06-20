using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class SumIf : HiddenValuesHandlingFunction
{
	private readonly ExpressionEvaluator _evaluator;

	public SumIf()
		: this(new ExpressionEvaluator())
	{
	}

	public SumIf(ExpressionEvaluator evaluator)
	{
		OfficeOpenXml.FormulaParsing.Utilities.Require.That(evaluator).Named("evaluator").IsNotNull();
		_evaluator = evaluator;
	}

	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 2);
		ExcelDataProvider.IRangeInfo rangeInfo = arguments.ElementAt(0).Value as ExcelDataProvider.IRangeInfo;
		string text = ((arguments.ElementAt(1).ValueFirst != null) ? ArgToString(arguments, 1) : null);
		double num = 0.0;
		if (rangeInfo == null)
		{
			object value = arguments.ElementAt(0).Value;
			if (text != null && _evaluator.Evaluate(value, text))
			{
				ExcelDataProvider.IRangeInfo source = arguments.ElementAt(2).Value as ExcelDataProvider.IRangeInfo;
				num = ((arguments.Count() > 2) ? source.First().ValueDouble : ConvertUtil.GetValueDouble(value, ignoreBool: true));
			}
		}
		else if (arguments.Count() > 2)
		{
			ExcelDataProvider.IRangeInfo sumRange = arguments.ElementAt(2).Value as ExcelDataProvider.IRangeInfo;
			num = CalculateWithSumRange(rangeInfo, text, sumRange, context);
		}
		else
		{
			num = CalculateSingleRange(rangeInfo, text, context);
		}
		return CreateResult(num, DataType.Decimal);
	}

	private double CalculateWithSumRange(ExcelDataProvider.IRangeInfo range, string criteria, ExcelDataProvider.IRangeInfo sumRange, ParsingContext context)
	{
		double num = 0.0;
		foreach (ExcelDataProvider.ICellInfo item in range)
		{
			if (criteria == null || !_evaluator.Evaluate(item.Value, criteria))
			{
				continue;
			}
			int num2 = item.Row - range.Address._fromRow;
			int num3 = item.Column - range.Address._fromCol;
			if (sumRange.Address._fromRow + num2 <= sumRange.Address._toRow && sumRange.Address._fromCol + num3 <= sumRange.Address._toCol)
			{
				object offset = sumRange.GetOffset(num2, num3);
				if (offset is ExcelErrorValue)
				{
					throw new ExcelErrorValueException((ExcelErrorValue)offset);
				}
				num += ConvertUtil.GetValueDouble(offset, ignoreBool: true);
			}
		}
		return num;
	}

	private double CalculateSingleRange(ExcelDataProvider.IRangeInfo range, string expression, ParsingContext context)
	{
		double num = 0.0;
		foreach (ExcelDataProvider.ICellInfo item in range)
		{
			if (expression != null && IsNumeric(item.Value) && _evaluator.Evaluate(item.Value, expression) && IsNumeric(item.Value))
			{
				if (item.IsExcelError)
				{
					throw new ExcelErrorValueException((ExcelErrorValue)item.Value);
				}
				num += item.ValueDouble;
			}
		}
		return num;
	}
}
