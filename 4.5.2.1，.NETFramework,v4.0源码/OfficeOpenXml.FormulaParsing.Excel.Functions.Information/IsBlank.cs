using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Information;

public class IsBlank : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		if (arguments == null || arguments.Count() == 0)
		{
			return CreateResult(true, DataType.Boolean);
		}
		bool flag = true;
		foreach (FunctionArgument argument in arguments)
		{
			if (argument.Value is ExcelDataProvider.IRangeInfo)
			{
				ExcelDataProvider.IRangeInfo rangeInfo = (ExcelDataProvider.IRangeInfo)argument.Value;
				if (rangeInfo.GetValue(rangeInfo.Address._fromRow, rangeInfo.Address._fromCol) != null)
				{
					flag = false;
				}
			}
			else if (argument.Value != null && argument.Value.ToString() != string.Empty)
			{
				flag = false;
				break;
			}
		}
		return CreateResult(flag, DataType.Boolean);
	}
}
