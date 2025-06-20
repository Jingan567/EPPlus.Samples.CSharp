using System;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class Sumsq : HiddenValuesHandlingFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		double num = 0.0;
		if (arguments != null)
		{
			foreach (FunctionArgument argument in arguments)
			{
				num += Calculate(argument, context);
			}
		}
		return CreateResult(num, DataType.Decimal);
	}

	private double Calculate(FunctionArgument arg, ParsingContext context, bool isInArray = false)
	{
		double num = 0.0;
		if (ShouldIgnore(arg))
		{
			return num;
		}
		if (arg.Value is IEnumerable<FunctionArgument>)
		{
			foreach (FunctionArgument item in (IEnumerable<FunctionArgument>)arg.Value)
			{
				num += Calculate(item, context, isInArray: true);
			}
		}
		else if (arg.Value is ExcelDataProvider.IRangeInfo rangeInfo)
		{
			foreach (ExcelDataProvider.ICellInfo item2 in rangeInfo)
			{
				if (!ShouldIgnore(item2, context))
				{
					CheckForAndHandleExcelError(item2);
					num += System.Math.Pow(item2.ValueDouble, 2.0);
				}
			}
		}
		else
		{
			CheckForAndHandleExcelError(arg);
			if (IsNumericString(arg.Value) && !isInArray)
			{
				return System.Math.Pow(ConvertUtil.GetValueDouble(arg.Value), 2.0);
			}
			bool ignoreBool = isInArray;
			num += System.Math.Pow(ConvertUtil.GetValueDouble(arg.Value, ignoreBool), 2.0);
		}
		return num;
	}
}
