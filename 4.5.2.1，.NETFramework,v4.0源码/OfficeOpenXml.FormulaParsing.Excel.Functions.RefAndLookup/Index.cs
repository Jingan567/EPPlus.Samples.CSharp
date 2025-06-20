using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

public class Index : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 2);
		FunctionArgument functionArgument = arguments.ElementAt(0);
		IEnumerable<FunctionArgument> enumerable = functionArgument.Value as IEnumerable<FunctionArgument>;
		CompileResultFactory compileResultFactory = new CompileResultFactory();
		if (enumerable != null)
		{
			int num = ArgToInt(arguments, 1);
			if (num > enumerable.Count())
			{
				throw new ExcelErrorValueException(eErrorType.Ref);
			}
			FunctionArgument functionArgument2 = enumerable.ElementAt(num - 1);
			return compileResultFactory.Create(functionArgument2.Value);
		}
		if (functionArgument.IsExcelRange)
		{
			int num2 = ArgToInt(arguments, 1);
			int num3 = ((arguments.Count() <= 2) ? 1 : ArgToInt(arguments, 2));
			ExcelDataProvider.IRangeInfo valueAsRangeInfo = functionArgument.ValueAsRangeInfo;
			if (num2 > valueAsRangeInfo.Address._toRow - valueAsRangeInfo.Address._fromRow + 1 || num3 > valueAsRangeInfo.Address._toCol - valueAsRangeInfo.Address._fromCol + 1)
			{
				ThrowExcelErrorValueException(eErrorType.Ref);
			}
			object offset = valueAsRangeInfo.GetOffset(num2 - 1, num3 - 1);
			return compileResultFactory.Create(offset);
		}
		throw new NotImplementedException();
	}
}
