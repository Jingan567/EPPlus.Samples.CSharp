using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Information;

public class IsError : ErrorHandlingFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		if (arguments == null || arguments.Count() == 0)
		{
			return CreateResult(false, DataType.Boolean);
		}
		foreach (FunctionArgument argument in arguments)
		{
			if (argument.Value is ExcelDataProvider.IRangeInfo)
			{
				ExcelDataProvider.IRangeInfo rangeInfo = (ExcelDataProvider.IRangeInfo)argument.Value;
				if (ExcelErrorValue.Values.IsErrorValue(rangeInfo.GetValue(rangeInfo.Address._fromRow, rangeInfo.Address._fromCol)))
				{
					return CreateResult(true, DataType.Boolean);
				}
			}
			else if (ExcelErrorValue.Values.IsErrorValue(argument.Value))
			{
				return CreateResult(true, DataType.Boolean);
			}
		}
		return CreateResult(false, DataType.Boolean);
	}

	public override CompileResult HandleError(string errorCode)
	{
		return CreateResult(true, DataType.Boolean);
	}
}
