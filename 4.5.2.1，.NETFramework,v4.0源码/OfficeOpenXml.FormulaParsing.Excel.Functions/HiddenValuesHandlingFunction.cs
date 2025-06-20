using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions;

public abstract class HiddenValuesHandlingFunction : ExcelFunction
{
	public bool IgnoreHiddenValues { get; set; }

	protected override IEnumerable<ExcelDoubleCellValue> ArgsToDoubleEnumerable(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		return ArgsToDoubleEnumerable(arguments, context, ignoreErrors: true);
	}

	protected IEnumerable<ExcelDoubleCellValue> ArgsToDoubleEnumerable(IEnumerable<FunctionArgument> arguments, ParsingContext context, bool ignoreErrors)
	{
		if (!arguments.Any())
		{
			return Enumerable.Empty<ExcelDoubleCellValue>();
		}
		if (IgnoreHiddenValues)
		{
			IEnumerable<FunctionArgument> arguments2 = arguments.Where((FunctionArgument x) => !x.ExcelStateFlagIsSet(ExcelCellState.HiddenCell));
			return base.ArgsToDoubleEnumerable(IgnoreHiddenValues, arguments2, context);
		}
		return base.ArgsToDoubleEnumerable(IgnoreHiddenValues, ignoreErrors, arguments, context);
	}

	protected bool ShouldIgnore(ExcelDataProvider.ICellInfo c, ParsingContext context)
	{
		return CellStateHelper.ShouldIgnore(IgnoreHiddenValues, c, context);
	}

	protected bool ShouldIgnore(FunctionArgument arg)
	{
		if (IgnoreHiddenValues && arg.ExcelStateFlagIsSet(ExcelCellState.HiddenCell))
		{
			return true;
		}
		return false;
	}
}
