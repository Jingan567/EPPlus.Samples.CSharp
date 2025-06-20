using System;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions;

public class DoubleEnumerableArgConverter : CollectionFlattener<ExcelDoubleCellValue>
{
	public virtual IEnumerable<ExcelDoubleCellValue> ConvertArgs(bool ignoreHidden, bool ignoreErrors, IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		return base.FuncArgsToFlatEnumerable(arguments, delegate(FunctionArgument arg, IList<ExcelDoubleCellValue> argList)
		{
			if (arg.IsExcelRange)
			{
				foreach (ExcelDataProvider.ICellInfo item3 in arg.ValueAsRangeInfo)
				{
					if (!ignoreErrors && item3.IsExcelError)
					{
						throw new ExcelErrorValueException(ExcelErrorValue.Parse(item3.Value.ToString()));
					}
					if (!CellStateHelper.ShouldIgnore(ignoreHidden, item3, context) && ConvertUtil.IsNumeric(item3.Value))
					{
						ExcelDoubleCellValue item = new ExcelDoubleCellValue(item3.ValueDouble, item3.Row);
						argList.Add(item);
					}
				}
				return;
			}
			if (!ignoreErrors && arg.ValueIsExcelError)
			{
				throw new ExcelErrorValueException(arg.ValueAsExcelErrorValue);
			}
			if (ConvertUtil.IsNumeric(arg.Value) && !CellStateHelper.ShouldIgnore(ignoreHidden, arg, context))
			{
				ExcelDoubleCellValue item2 = new ExcelDoubleCellValue(ConvertUtil.GetValueDouble(arg.Value));
				argList.Add(item2);
			}
		});
	}

	public virtual IEnumerable<ExcelDoubleCellValue> ConvertArgsIncludingOtherTypes(IEnumerable<FunctionArgument> arguments)
	{
		return base.FuncArgsToFlatEnumerable(arguments, delegate(FunctionArgument arg, IList<ExcelDoubleCellValue> argList)
		{
			if (arg.Value is ExcelDataProvider.IRangeInfo)
			{
				foreach (ExcelDataProvider.ICellInfo item2 in (ExcelDataProvider.IRangeInfo)arg.Value)
				{
					ExcelDoubleCellValue item = new ExcelDoubleCellValue(item2.ValueDoubleLogical, item2.Row);
					argList.Add(item);
				}
				return;
			}
			if (arg.Value is double || arg.Value is int || arg.Value is bool)
			{
				argList.Add(Convert.ToDouble(arg.Value));
			}
			else if (arg.Value is string)
			{
				argList.Add(0.0);
			}
		});
	}
}
