using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

public class Columns : LookupFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		ExcelDataProvider.IRangeInfo valueAsRangeInfo = arguments.ElementAt(0).ValueAsRangeInfo;
		if (valueAsRangeInfo != null)
		{
			return CreateResult(valueAsRangeInfo.Address._toCol - valueAsRangeInfo.Address._fromCol + 1, DataType.Integer);
		}
		string text = ArgToString(arguments, 0);
		if (ExcelAddressUtil.IsValidAddress(text))
		{
			RangeAddress rangeAddress = new RangeAddressFactory(context.ExcelDataProvider).Create(text);
			return CreateResult(rangeAddress.ToCol - rangeAddress.FromCol + 1, DataType.Integer);
		}
		throw new ArgumentException("Invalid range supplied");
	}
}
