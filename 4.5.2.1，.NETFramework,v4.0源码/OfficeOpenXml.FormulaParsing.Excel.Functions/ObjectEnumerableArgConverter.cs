using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions;

public class ObjectEnumerableArgConverter : CollectionFlattener<object>
{
	public virtual IEnumerable<object> ConvertArgs(bool ignoreHidden, IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		return base.FuncArgsToFlatEnumerable(arguments, delegate(FunctionArgument arg, IList<object> argList)
		{
			if (arg.Value is ExcelDataProvider.IRangeInfo)
			{
				foreach (ExcelDataProvider.ICellInfo item in (ExcelDataProvider.IRangeInfo)arg.Value)
				{
					if (!CellStateHelper.ShouldIgnore(ignoreHidden, item, context))
					{
						argList.Add(item.Value);
					}
				}
				return;
			}
			argList.Add(arg.Value);
		});
	}
}
