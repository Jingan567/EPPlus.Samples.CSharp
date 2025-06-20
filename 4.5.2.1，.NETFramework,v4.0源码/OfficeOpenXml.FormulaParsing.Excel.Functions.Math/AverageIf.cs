using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class AverageIf : HiddenValuesHandlingFunction
{
	private readonly ExpressionEvaluator _expressionEvaluator;

	public AverageIf()
		: this(new ExpressionEvaluator())
	{
	}

	public AverageIf(ExpressionEvaluator evaluator)
	{
		OfficeOpenXml.FormulaParsing.Utilities.Require.That(evaluator).Named("evaluator").IsNotNull();
		_expressionEvaluator = evaluator;
	}

	private bool Evaluate(object obj, string expression)
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

	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 2);
		ExcelDataProvider.IRangeInfo rangeInfo = arguments.ElementAt(0).Value as ExcelDataProvider.IRangeInfo;
		string text = ((arguments.ElementAt(1).ValueFirst != null) ? ArgToString(arguments, 1) : null);
		double num = 0.0;
		if (rangeInfo == null)
		{
			object value = arguments.ElementAt(0).Value;
			if (text == null || !Evaluate(value, text))
			{
				throw new ExcelErrorValueException(eErrorType.Div0);
			}
			ExcelDataProvider.IRangeInfo source = arguments.ElementAt(2).Value as ExcelDataProvider.IRangeInfo;
			num = ((arguments.Count() > 2) ? source.First().ValueDouble : ConvertUtil.GetValueDouble(value, ignoreBool: true));
		}
		else if (arguments.Count() > 2)
		{
			ExcelDataProvider.IRangeInfo sumRange = arguments.ElementAt(2).Value as ExcelDataProvider.IRangeInfo;
			num = CalculateWithLookupRange(rangeInfo, text, sumRange, context);
		}
		else
		{
			num = CalculateSingleRange(rangeInfo, text, context);
		}
		return CreateResult(num, DataType.Decimal);
	}

	private double CalculateWithLookupRange(ExcelDataProvider.IRangeInfo range, string criteria, ExcelDataProvider.IRangeInfo sumRange, ParsingContext context)
	{
		double num = 0.0;
		int num2 = 0;
		foreach (ExcelDataProvider.ICellInfo item in range)
		{
			if (criteria == null || !Evaluate(item.Value, criteria))
			{
				continue;
			}
			int num3 = item.Row - range.Address._fromRow;
			int num4 = item.Column - range.Address._fromCol;
			if (sumRange.Address._fromRow + num3 <= sumRange.Address._toRow && sumRange.Address._fromCol + num4 <= sumRange.Address._toCol)
			{
				object offset = sumRange.GetOffset(num3, num4);
				if (offset is ExcelErrorValue)
				{
					throw new ExcelErrorValueException((ExcelErrorValue)offset);
				}
				num2++;
				num += ConvertUtil.GetValueDouble(offset, ignoreBool: true);
			}
		}
		return Divide(num, num2);
	}

	private double CalculateSingleRange(ExcelDataProvider.IRangeInfo range, string expression, ParsingContext context)
	{
		double num = 0.0;
		int num2 = 0;
		foreach (ExcelDataProvider.ICellInfo item in range)
		{
			if (expression != null && IsNumeric(item.Value) && Evaluate(item.Value, expression))
			{
				if (item.IsExcelError)
				{
					throw new ExcelErrorValueException((ExcelErrorValue)item.Value);
				}
				num += item.ValueDouble;
				num2++;
			}
		}
		return Divide(num, num2);
	}
}
