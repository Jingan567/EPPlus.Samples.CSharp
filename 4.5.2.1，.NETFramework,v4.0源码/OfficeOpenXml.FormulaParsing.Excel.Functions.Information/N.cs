using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Information;

public class N : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		object firstValue = GetFirstValue(arguments);
		if (firstValue is bool)
		{
			double num = (((bool)firstValue) ? 1.0 : 0.0);
			return CreateResult(num, DataType.Decimal);
		}
		if (IsNumeric(firstValue))
		{
			double valueDouble = ConvertUtil.GetValueDouble(firstValue);
			return CreateResult(valueDouble, DataType.Decimal);
		}
		if (firstValue is string)
		{
			return CreateResult(0.0, DataType.Decimal);
		}
		if (firstValue is ExcelErrorValue)
		{
			return CreateResult(firstValue, DataType.ExcelError);
		}
		throw new ExcelErrorValueException(eErrorType.Value);
	}
}
