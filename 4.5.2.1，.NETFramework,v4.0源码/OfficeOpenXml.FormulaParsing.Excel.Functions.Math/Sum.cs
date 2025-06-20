using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class Sum : HiddenValuesHandlingFunction
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

	private double Calculate(FunctionArgument arg, ParsingContext context)
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
				num += Calculate(item, context);
			}
		}
		else if (arg.Value is ExcelDataProvider.IRangeInfo)
		{
			foreach (ExcelDataProvider.ICellInfo item2 in (ExcelDataProvider.IRangeInfo)arg.Value)
			{
				if (!ShouldIgnore(item2, context))
				{
					CheckForAndHandleExcelError(item2);
					num += item2.ValueDouble;
				}
			}
		}
		else
		{
			CheckForAndHandleExcelError(arg);
			num += ConvertUtil.GetValueDouble(arg.Value, ignoreBool: true);
		}
		return num;
	}
}
