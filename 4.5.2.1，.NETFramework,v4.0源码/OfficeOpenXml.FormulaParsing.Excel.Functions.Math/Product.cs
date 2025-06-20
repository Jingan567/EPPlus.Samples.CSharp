using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class Product : HiddenValuesHandlingFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		double num = 0.0;
		int num2 = 0;
		while (AreEqual(num, 0.0) && num2 < arguments.Count())
		{
			num = CalculateFirstItem(arguments, num2++, context);
		}
		num = CalculateCollection(arguments.Skip(num2), num, delegate(FunctionArgument arg, double current)
		{
			if (ShouldIgnore(arg))
			{
				return current;
			}
			if (arg.ValueIsExcelError)
			{
				ThrowExcelErrorValueException(arg.ValueAsExcelErrorValue.Type);
			}
			if (arg.IsExcelRange)
			{
				foreach (ExcelDataProvider.ICellInfo item in arg.ValueAsRangeInfo)
				{
					if (ShouldIgnore(item, context))
					{
						return current;
					}
					current *= item.ValueDouble;
				}
				return current;
			}
			object value = arg.Value;
			if (value != null && IsNumeric(value))
			{
				double num3 = Convert.ToDouble(value);
				current *= num3;
			}
			return current;
		});
		return CreateResult(num, DataType.Decimal);
	}

	private double CalculateFirstItem(IEnumerable<FunctionArgument> arguments, int index, ParsingContext context)
	{
		FunctionArgument item = arguments.ElementAt(index);
		List<FunctionArgument> arguments2 = new List<FunctionArgument> { item };
		IEnumerable<ExcelDoubleCellValue> enumerable = ArgsToDoubleEnumerable(ignoreHiddenCells: false, ignoreErrors: false, arguments2, context);
		double num = 0.0;
		foreach (ExcelDoubleCellValue item2 in enumerable)
		{
			num = ((num != 0.0 || !((double)item2 > 0.0)) ? (num * (double)item2) : ((double)item2));
		}
		return num;
	}
}
