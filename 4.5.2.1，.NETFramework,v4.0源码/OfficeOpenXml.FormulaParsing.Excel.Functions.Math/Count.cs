using System;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class Count : HiddenValuesHandlingFunction
{
	private enum ItemContext
	{
		InRange,
		InArray,
		SingleArg
	}

	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		double nItems = 0.0;
		Calculate(arguments, ref nItems, context, ItemContext.SingleArg);
		return CreateResult(nItems, DataType.Integer);
	}

	private void Calculate(IEnumerable<FunctionArgument> items, ref double nItems, ParsingContext context, ItemContext itemContext)
	{
		foreach (FunctionArgument item in items)
		{
			if (item.Value is ExcelDataProvider.IRangeInfo rangeInfo)
			{
				foreach (ExcelDataProvider.ICellInfo item2 in rangeInfo)
				{
					_CheckForAndHandleExcelError(item2, context);
					if (!ShouldIgnore(item2, context) && ShouldCount(item2.Value, ItemContext.InRange))
					{
						nItems += 1.0;
					}
				}
			}
			else if (item.Value is IEnumerable<FunctionArgument> items2)
			{
				Calculate(items2, ref nItems, context, ItemContext.InArray);
			}
			else
			{
				_CheckForAndHandleExcelError(item, context);
				if (!ShouldIgnore(item) && ShouldCount(item.Value, itemContext))
				{
					nItems += 1.0;
				}
			}
		}
	}

	private void _CheckForAndHandleExcelError(FunctionArgument arg, ParsingContext context)
	{
	}

	private void _CheckForAndHandleExcelError(ExcelDataProvider.ICellInfo cell, ParsingContext context)
	{
	}

	private bool ShouldCount(object value, ItemContext context)
	{
		switch (context)
		{
		case ItemContext.SingleArg:
			if (!IsNumeric(value))
			{
				return IsNumericString(value);
			}
			return true;
		case ItemContext.InRange:
			return IsNumeric(value);
		case ItemContext.InArray:
			if (!IsNumeric(value))
			{
				return IsNumericString(value);
			}
			return true;
		default:
			throw new ArgumentException("Unknown ItemContext:" + context);
		}
	}
}
