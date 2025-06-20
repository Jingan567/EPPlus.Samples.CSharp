using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class CountIf : ExcelFunction
{
	private readonly ExpressionEvaluator _expressionEvaluator;

	public CountIf()
		: this(new ExpressionEvaluator())
	{
	}

	public CountIf(ExpressionEvaluator evaluator)
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
		FunctionArgument[] array = (arguments as FunctionArgument[]) ?? arguments.ToArray();
		ValidateArguments(array, 2);
		FunctionArgument functionArgument = array.ElementAt(0);
		string text = ((array.ElementAt(1).ValueFirst != null) ? ArgToString(array, 1) : null);
		double num = 0.0;
		if (functionArgument.IsExcelRange)
		{
			ExcelDataProvider.IRangeInfo valueAsRangeInfo = functionArgument.ValueAsRangeInfo;
			for (int i = valueAsRangeInfo.Address.Start.Row; i < valueAsRangeInfo.Address.End.Row + 1; i++)
			{
				for (int j = valueAsRangeInfo.Address.Start.Column; j < valueAsRangeInfo.Address.End.Column + 1; j++)
				{
					if (text != null && Evaluate(valueAsRangeInfo.Worksheet.GetValue(i, j), text))
					{
						num += 1.0;
					}
				}
			}
		}
		else if (functionArgument.Value is IEnumerable<FunctionArgument>)
		{
			foreach (FunctionArgument item in (IEnumerable<FunctionArgument>)functionArgument.Value)
			{
				if (Evaluate(item.Value, text))
				{
					num += 1.0;
				}
			}
		}
		else if (Evaluate(functionArgument.Value, text))
		{
			num += 1.0;
		}
		return CreateResult(num, DataType.Integer);
	}
}
