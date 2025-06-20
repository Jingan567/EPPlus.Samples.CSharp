using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class SumProduct : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		double num = 0.0;
		List<List<double>> list = new List<List<double>>();
		foreach (FunctionArgument argument in arguments)
		{
			list.Add(new List<double>());
			List<double> currentResult = list.Last();
			if (argument.Value is IEnumerable<FunctionArgument>)
			{
				foreach (FunctionArgument item in (IEnumerable<FunctionArgument>)argument.Value)
				{
					AddValue(item.Value, currentResult);
				}
			}
			else if (argument.Value is FunctionArgument)
			{
				AddValue(argument.Value, currentResult);
			}
			else if (argument.IsExcelRange)
			{
				ExcelDataProvider.IRangeInfo valueAsRangeInfo = argument.ValueAsRangeInfo;
				for (int i = valueAsRangeInfo.Address._fromCol; i <= valueAsRangeInfo.Address._toCol; i++)
				{
					for (int j = valueAsRangeInfo.Address._fromRow; j <= valueAsRangeInfo.Address._toRow; j++)
					{
						AddValue(valueAsRangeInfo.GetValue(j, i), currentResult);
					}
				}
			}
			else if (IsNumeric(argument.Value))
			{
				AddValue(argument.Value, currentResult);
			}
		}
		int count = list.First().Count;
		foreach (List<double> item2 in list)
		{
			if (item2.Count != count)
			{
				throw new ExcelErrorValueException(ExcelErrorValue.Create(eErrorType.Value));
			}
		}
		for (int k = 0; k < count; k++)
		{
			double num2 = 1.0;
			for (int l = 0; l < list.Count; l++)
			{
				num2 *= list[l][k];
			}
			num += num2;
		}
		return CreateResult(num, DataType.Decimal);
	}

	private void AddValue(object convertVal, List<double> currentResult)
	{
		if (IsNumeric(convertVal))
		{
			currentResult.Add(Convert.ToDouble(convertVal));
			return;
		}
		if (convertVal is ExcelErrorValue)
		{
			throw new ExcelErrorValueException((ExcelErrorValue)convertVal);
		}
		currentResult.Add(0.0);
	}
}
