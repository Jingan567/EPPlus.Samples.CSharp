using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

public class Address : ExcelFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 2);
		int row = ArgToInt(arguments, 0);
		int col = ArgToInt(arguments, 1);
		ThrowExcelErrorValueExceptionIf(() => row < 0 && col < 0, eErrorType.Value);
		ExcelReferenceType referenceType = ExcelReferenceType.AbsoluteRowAndColumn;
		string text = string.Empty;
		if (arguments.Count() > 2)
		{
			int arg3 = ArgToInt(arguments, 2);
			ThrowExcelErrorValueExceptionIf(() => arg3 < 1 || arg3 > 4, eErrorType.Value);
			referenceType = (ExcelReferenceType)ArgToInt(arguments, 2);
		}
		if (arguments.Count() > 3)
		{
			object value = arguments.ElementAt(3).Value;
			if (value is bool && !(bool)value)
			{
				throw new InvalidOperationException("Excelformulaparser does not support the R1C1 format!");
			}
		}
		if (arguments.Count() > 4)
		{
			object value2 = arguments.ElementAt(4).Value;
			if (value2 is string && !string.IsNullOrEmpty(value2.ToString()))
			{
				text = string.Concat(value2, "!");
			}
		}
		IndexToAddressTranslator indexToAddressTranslator = new IndexToAddressTranslator(context.ExcelDataProvider, referenceType);
		return CreateResult(text + indexToAddressTranslator.ToAddress(col, row), DataType.ExcelAddress);
	}
}
