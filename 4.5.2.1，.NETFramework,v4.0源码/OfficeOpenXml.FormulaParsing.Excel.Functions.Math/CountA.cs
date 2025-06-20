using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class CountA : HiddenValuesHandlingFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		double nItems = 0.0;
		Calculate(arguments, context, ref nItems);
		return CreateResult(nItems, DataType.Integer);
	}

	private void Calculate(IEnumerable<FunctionArgument> items, ParsingContext context, ref double nItems)
	{
		foreach (FunctionArgument item in items)
		{
			if (item.Value is ExcelDataProvider.IRangeInfo rangeInfo)
			{
				foreach (ExcelDataProvider.ICellInfo item2 in rangeInfo)
				{
					_CheckForAndHandleExcelError(item2, context);
					if (!ShouldIgnore(item2, context) && ShouldCount(item2.Value))
					{
						nItems += 1.0;
					}
				}
			}
			else if (item.Value is IEnumerable<FunctionArgument>)
			{
				Calculate((IEnumerable<FunctionArgument>)item.Value, context, ref nItems);
			}
			else
			{
				_CheckForAndHandleExcelError(item, context);
				if (!ShouldIgnore(item) && ShouldCount(item.Value))
				{
					nItems += 1.0;
				}
			}
		}
	}

	private void _CheckForAndHandleExcelError(FunctionArgument arg, ParsingContext context)
	{
		if (context.Scopes.Current.IsSubtotal)
		{
			CheckForAndHandleExcelError(arg);
		}
	}

	private void _CheckForAndHandleExcelError(ExcelDataProvider.ICellInfo cell, ParsingContext context)
	{
		if (context.Scopes.Current.IsSubtotal)
		{
			CheckForAndHandleExcelError(cell);
		}
	}

	private bool ShouldCount(object value)
	{
		if (value == null)
		{
			return false;
		}
		return !string.IsNullOrEmpty(value.ToString());
	}
}
