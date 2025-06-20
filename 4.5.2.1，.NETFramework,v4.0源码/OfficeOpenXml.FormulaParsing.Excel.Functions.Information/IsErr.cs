using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Information;

public class IsErr : ErrorHandlingFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		CompileResult compileResult = new IsError().Execute(arguments, context);
		if ((bool)compileResult.Result)
		{
			object firstValue = GetFirstValue(arguments);
			if (firstValue is ExcelDataProvider.IRangeInfo)
			{
				ExcelDataProvider.IRangeInfo rangeInfo = (ExcelDataProvider.IRangeInfo)firstValue;
				if (rangeInfo.GetValue(rangeInfo.Address._fromRow, rangeInfo.Address._fromCol) is ExcelErrorValue { Type: eErrorType.NA })
				{
					return CreateResult(false, DataType.Boolean);
				}
			}
			else if (firstValue is ExcelErrorValue && ((ExcelErrorValue)firstValue).Type == eErrorType.NA)
			{
				return CreateResult(false, DataType.Boolean);
			}
		}
		return compileResult;
	}

	public override CompileResult HandleError(string errorCode)
	{
		return CreateResult(true, DataType.Boolean);
	}
}
