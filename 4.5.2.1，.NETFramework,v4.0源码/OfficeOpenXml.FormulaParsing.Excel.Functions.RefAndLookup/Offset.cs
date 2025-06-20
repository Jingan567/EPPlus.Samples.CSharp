using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

public class Offset : LookupFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		FunctionArgument[] array = (arguments as FunctionArgument[]) ?? arguments.ToArray();
		ValidateArguments(array, 3);
		string address = ArgToAddress(array, 0);
		int num = ArgToInt(array, 1);
		int num2 = ArgToInt(array, 2);
		int width = 0;
		int height = 0;
		if (array.Length > 3)
		{
			height = ArgToInt(array, 3);
			ThrowExcelErrorValueExceptionIf(() => height == 0, eErrorType.Ref);
		}
		if (array.Length > 4)
		{
			width = ArgToInt(array, 4);
			ThrowExcelErrorValueExceptionIf(() => width == 0, eErrorType.Ref);
		}
		string worksheet = context.Scopes.Current.Address.Worksheet;
		ExcelAddressBase address2 = context.ExcelDataProvider.GetRange(worksheet, address).Address;
		int num3 = address2._fromRow + num;
		int num4 = address2._fromCol + num2;
		int toRow = ((height != 0) ? (address2._fromRow + height - 1) : address2._toRow) + num;
		int toCol = ((width != 0) ? (address2._fromCol + width - 1) : address2._toCol) + num2;
		ExcelDataProvider.IRangeInfo range = context.ExcelDataProvider.GetRange(address2.WorkSheet, num3, num4, toRow, toCol);
		if (!range.IsMulti)
		{
			if (range.IsEmpty)
			{
				return CompileResult.Empty;
			}
			object value = range.GetValue(num3, num4);
			if (IsNumeric(value))
			{
				return CreateResult(value, DataType.Decimal);
			}
			if (value is ExcelErrorValue)
			{
				return CreateResult(value, DataType.ExcelError);
			}
			return CreateResult(value, DataType.String);
		}
		return CreateResult(range, DataType.Enumerable);
	}
}
