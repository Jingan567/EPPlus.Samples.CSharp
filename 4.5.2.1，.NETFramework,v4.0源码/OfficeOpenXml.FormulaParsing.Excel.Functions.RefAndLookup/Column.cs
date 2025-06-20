using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

public class Column : LookupFunction
{
	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		if (arguments == null || arguments.Count() == 0)
		{
			return CreateResult(context.Scopes.Current.Address.FromCol, DataType.Integer);
		}
		string text = ArgToString(arguments, 0);
		if (!ExcelAddressUtil.IsValidAddress(text))
		{
			throw new ArgumentException("An invalid argument was supplied");
		}
		RangeAddress rangeAddress = new RangeAddressFactory(context.ExcelDataProvider).Create(text);
		return CreateResult(rangeAddress.FromCol, DataType.Integer);
	}
}
