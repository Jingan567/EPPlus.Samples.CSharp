using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class Rank : ExcelFunction
{
	private bool _isAvg;

	public Rank(bool isAvg = false)
	{
		_isAvg = isAvg;
	}

	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 2);
		double num = ArgToDecimal(arguments, 0);
		FunctionArgument functionArgument = arguments.ElementAt(1);
		bool flag = false;
		if (arguments.Count() > 2)
		{
			flag = ArgToBool(arguments, 2);
		}
		List<double> list = new List<double>();
		foreach (ExcelDataProvider.ICellInfo item in functionArgument.ValueAsRangeInfo)
		{
			double valueDouble = ConvertUtil.GetValueDouble(item.Value, ignoreBool: false, retNaN: true);
			if (!double.IsNaN(valueDouble))
			{
				list.Add(valueDouble);
			}
		}
		list.Sort();
		double num2;
		if (flag)
		{
			num2 = list.IndexOf(num) + 1;
			if (_isAvg)
			{
				int i;
				for (i = Convert.ToInt32(num2); list.Count > i && list[i] == num; i++)
				{
				}
				if ((double)i > num2)
				{
					num2 += ((double)i - num2) / 2.0;
				}
			}
		}
		else
		{
			num2 = list.LastIndexOf(num);
			if (_isAvg)
			{
				int num3 = Convert.ToInt32(num2) - 1;
				while (0 <= num3 && list[num3] == num)
				{
					num3--;
				}
				if ((double)(num3 + 1) < num2)
				{
					num2 -= (num2 - (double)num3 - 1.0) / 2.0;
				}
			}
			num2 = (double)list.Count - num2;
		}
		if (num2 <= 0.0 || num2 > (double)list.Count)
		{
			return new CompileResult(ExcelErrorValue.Create(eErrorType.NA), DataType.ExcelError);
		}
		return CreateResult(num2, DataType.Decimal);
	}
}
