using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class Small : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 2);
		FunctionArgument item = arguments.ElementAt(0);
		int index = ArgToInt(arguments, 1) - 1;
		IEnumerable<ExcelDoubleCellValue> values = ArgsToDoubleEnumerable(new List<FunctionArgument> { item }, context);
		ThrowExcelErrorValueExceptionIf(() => index < 0 || index >= values.Count(), eErrorType.Num);
		ExcelDoubleCellValue excelDoubleCellValue = values.OrderBy((ExcelDoubleCellValue x) => x).ElementAt(index);
		return CreateResult(excelDoubleCellValue, DataType.Decimal);
	}
}
