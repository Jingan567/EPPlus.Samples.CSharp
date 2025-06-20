using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class Average : HiddenValuesHandlingFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1, eErrorType.Div0);
		double nValues = 0.0;
		double retVal = 0.0;
		foreach (FunctionArgument argument in arguments)
		{
			Calculate(argument, context, ref retVal, ref nValues);
		}
		return CreateResult(Divide(retVal, nValues), DataType.Decimal);
	}

	private void Calculate(FunctionArgument arg, ParsingContext context, ref double retVal, ref double nValues, bool isInArray = false)
	{
		if (ShouldIgnore(arg))
		{
			return;
		}
		if (arg.Value is IEnumerable<FunctionArgument>)
		{
			foreach (FunctionArgument item in (IEnumerable<FunctionArgument>)arg.Value)
			{
				Calculate(item, context, ref retVal, ref nValues, isInArray: true);
			}
		}
		else if (arg.IsExcelRange)
		{
			foreach (ExcelDataProvider.ICellInfo item2 in arg.ValueAsRangeInfo)
			{
				if (!ShouldIgnore(item2, context))
				{
					CheckForAndHandleExcelError(item2);
					if (IsNumeric(item2.Value) && !(item2.Value is bool))
					{
						nValues += 1.0;
						retVal += item2.ValueDouble;
					}
				}
			}
		}
		else
		{
			double? numericValue = GetNumericValue(arg.Value, isInArray);
			if (numericValue.HasValue)
			{
				nValues += 1.0;
				retVal += numericValue.Value;
			}
			else if (arg.Value is string && !isInArray)
			{
				ThrowExcelErrorValueException(eErrorType.Value);
			}
		}
		CheckForAndHandleExcelError(arg);
	}

	private double? GetNumericValue(object obj, bool isInArray)
	{
		if (IsNumeric(obj) && !(obj is bool))
		{
			return ConvertUtil.GetValueDouble(obj);
		}
		if (!isInArray)
		{
			if (obj is bool)
			{
				return ConvertUtil.GetValueDouble(obj);
			}
			if (ConvertUtil.TryParseNumericString(obj, out var result))
			{
				return result;
			}
			if (ConvertUtil.TryParseDateString(obj, out var result2))
			{
				return result2.ToOADate();
			}
		}
		return null;
	}
}
